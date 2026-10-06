#!/usr/bin/env pwsh

# Required parameters:
# @raycast.schemaVersion 1
# @raycast.title Codex Status
# @raycast.mode compact

# Optional parameters:
# @raycast.icon 📊
# @raycast.packageName Developer Tools
# @raycast.description Checks the running status of Win-Codex MCP server and Cloudflare tunnel.
# @raycast.author Deepak Bhattarai

$statuses = @()

# 1. Check Win-Codex (Port 6000)
$conn = Get-NetTCPConnection -LocalPort 6000 -ErrorAction SilentlyContinue | Select-Object -First 1
if ($conn) {
    $proc = Get-Process -Id $conn.OwningProcess -ErrorAction SilentlyContinue
    $statuses += "Codex: ONLINE (PID $($proc.Id), Port 6000)"
} else {
    $statuses += "Codex: OFFLINE"
}

# 2. Check Cloudflare Tunnel
$cfProc = Get-Process -Name "cloudflared" -ErrorAction SilentlyContinue | Select-Object -First 1
if ($cfProc) {
    $statuses += "Cloudflare: ONLINE (PID $($cfProc.Id))"
} else {
    $statuses += "Cloudflare: OFFLINE"
}

Write-Output ($statuses -join " | ")
