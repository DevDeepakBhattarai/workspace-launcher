#!/usr/bin/env pwsh

# Required parameters:
# @raycast.schemaVersion 1
# @raycast.title Restart Codex
# @raycast.mode compact

# Optional parameters:
# @raycast.icon 🔄
# @raycast.packageName Developer Tools
# @raycast.description Restarts both local Win-Codex MCP server and Cloudflare tunnel.
# @raycast.author Deepak Bhattarai

$stopScript = Join-Path $PSScriptRoot "stop-codex.ps1"
$startScript = Join-Path $PSScriptRoot "start-codex.ps1"

if (Test-Path $stopScript) {
    & $stopScript | Out-Null
}

Start-Sleep -Seconds 1

if (Test-Path $startScript) {
    & $startScript
} else {
    Write-Output "Start script not found"
}
