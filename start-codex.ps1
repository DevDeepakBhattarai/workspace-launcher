#!/usr/bin/env pwsh

# Required parameters:
# @raycast.schemaVersion 1
# @raycast.title Start Codex
# @raycast.mode compact

# Optional parameters:
# @raycast.icon 🚀
# @raycast.packageName Developer Tools
# @raycast.description Starts the local Win-Codex MCP server (port 6000) and Cloudflare tunnel in the background.
# @raycast.author Deepak Bhattarai

$wincodexDir = "D:\Coding\Experiments\local-windows-control-mcp"
$logDir = Join-Path $wincodexDir ".data"
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

$wincodexLog = Join-Path $logDir "server.log"
$wincodexErr = Join-Path $logDir "server-err.log"

$cloudflaredLogDir = "$env:USERPROFILE\.cloudflared"
if (-not (Test-Path $cloudflaredLogDir)) {
    New-Item -ItemType Directory -Path $cloudflaredLogDir -Force | Out-Null
}
$cloudflaredLog = Join-Path $cloudflaredLogDir "cloudflared.log"
$cloudflaredErr = Join-Path $cloudflaredLogDir "cloudflared-err.log"

$messages = @()

# 1. Start Win-Codex Server
$conn = Get-NetTCPConnection -LocalPort 6000 -ErrorAction SilentlyContinue | Select-Object -First 1
if ($conn) {
    $proc = Get-Process -Id $conn.OwningProcess -ErrorAction SilentlyContinue
    $messages += "Codex: already running (PID $($proc.Id))"
} else {
    $distPath = Join-Path $wincodexDir "dist\server.js"
    if (-not (Test-Path $distPath)) {
        Start-Process -FilePath "pnpm.cmd" -ArgumentList "build" -WorkingDirectory $wincodexDir -Wait -WindowStyle Hidden
    }

    Start-Process -FilePath "node.exe" `
        -ArgumentList "dist/server.js" `
        -WorkingDirectory $wincodexDir `
        -RedirectStandardOutput $wincodexLog `
        -RedirectStandardError $wincodexErr `
        -WindowStyle Hidden

    # Verify startup
    $started = $false
    for ($i = 0; $i -lt 15; $i++) {
        Start-Sleep -Milliseconds 200
        $conn = Get-NetTCPConnection -LocalPort 6000 -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($conn) {
            $started = $true
            break
        }
    }

    if ($started) {
        $messages += "Codex: started on port 6000"
    } else {
        $messages += "Codex: launched (check .data/server.log)"
    }
}

# 2. Start Cloudflare Tunnel
$cfProc = Get-Process -Name "cloudflared" -ErrorAction SilentlyContinue | Select-Object -First 1
if ($cfProc) {
    $messages += "Cloudflare: already running (PID $($cfProc.Id))"
} else {
    $cfExe = "cloudflared.exe"
    $cfCmd = Get-Command $cfExe -ErrorAction SilentlyContinue
    if (-not $cfCmd) {
        if (Test-Path "C:\Program Files (x86)\cloudflared\cloudflared.exe") {
            $cfExe = "C:\Program Files (x86)\cloudflared\cloudflared.exe"
        } elseif (Test-Path "C:\Program Files\cloudflared\cloudflared.exe") {
            $cfExe = "C:\Program Files\cloudflared\cloudflared.exe"
        }
    }

    Start-Process -FilePath $cfExe `
        -ArgumentList "tunnel run" `
        -RedirectStandardOutput $cloudflaredLog `
        -RedirectStandardError $cloudflaredErr `
        -WindowStyle Hidden

    Start-Sleep -Milliseconds 500
    $cfProc = Get-Process -Name "cloudflared" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($cfProc) {
        $messages += "Cloudflare: tunnel started (PID $($cfProc.Id))"
    } else {
        $messages += "Cloudflare: tunnel launched"
    }
}

Write-Output ($messages -join " | ")
