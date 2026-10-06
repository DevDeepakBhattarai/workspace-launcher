$ErrorActionPreference = 'Stop'

$packages = @(
    '@openai/codex@latest',
    '@anthropic-ai/claude-code@latest',
    'opencode-ai@latest'
)

$minimumReleaseAgeBefore = (& pnpm config get minimumReleaseAge).Trim()
$previousOverride = $env:PNPM_CONFIG_MINIMUM_RELEASE_AGE

try {
    # Ignore minimumReleaseAge only for this update command and its dependency tree.
    # The user's global pnpm minimumReleaseAge setting is not changed.
    $env:PNPM_CONFIG_MINIMUM_RELEASE_AGE = '0'

    Write-Host 'Updating Codex, Claude Code, and OpenCode to registry latest...'
    & pnpm add -g @packages
    if ($LASTEXITCODE -ne 0) {
        throw "pnpm exited with code $LASTEXITCODE"
    }
}
finally {
    if ($null -eq $previousOverride) {
        Remove-Item Env:PNPM_CONFIG_MINIMUM_RELEASE_AGE -ErrorAction SilentlyContinue
    }
    else {
        $env:PNPM_CONFIG_MINIMUM_RELEASE_AGE = $previousOverride
    }
}

# T3 Code checks %PNPM_HOME% before the configured global bin directory on Windows.
# Remove stale launchers left in the old PNPM_HOME root so T3 resolves the current shims.
$pnpmHome = [Environment]::GetEnvironmentVariable('PNPM_HOME', 'User')
if (-not $pnpmHome) {
    $pnpmHome = $env:PNPM_HOME
}

$globalBinDir = (& pnpm config get globalBinDir).Trim()
if ($pnpmHome -and $globalBinDir -and
    -not [string]::Equals(
        [IO.Path]::GetFullPath($pnpmHome).TrimEnd('\'),
        [IO.Path]::GetFullPath($globalBinDir).TrimEnd('\'),
        [StringComparison]::OrdinalIgnoreCase
    )) {
    foreach ($name in @('codex', 'claude', 'opencode')) {
        foreach ($suffix in @('.CMD', '.cmd', '.ps1', '')) {
            $staleShim = Join-Path $pnpmHome ($name + $suffix)
            if (Test-Path -LiteralPath $staleShim -PathType Leaf) {
                Remove-Item -LiteralPath $staleShim -Force
                Write-Host "Removed stale launcher: $staleShim"
            }
        }
    }
}

$minimumReleaseAgeAfter = (& pnpm config get minimumReleaseAge).Trim()
if ($minimumReleaseAgeAfter -ne $minimumReleaseAgeBefore) {
    throw "pnpm minimumReleaseAge changed unexpectedly from '$minimumReleaseAgeBefore' to '$minimumReleaseAgeAfter'."
}

Write-Host ''
Write-Host "pnpm minimumReleaseAge preserved: $minimumReleaseAgeAfter minutes"
Write-Host 'Installed versions:'
& codex --version
if ($LASTEXITCODE -ne 0) { throw 'codex --version failed' }
& claude --version
if ($LASTEXITCODE -ne 0) { throw 'claude --version failed' }
& opencode --version
if ($LASTEXITCODE -ne 0) { throw 'opencode --version failed' }
