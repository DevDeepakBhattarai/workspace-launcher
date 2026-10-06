#!/usr/bin/env pwsh

# Required parameters:
# @raycast.schemaVersion 1
# @raycast.title Stop Codex
# @raycast.mode compact

# Optional parameters:
# @raycast.icon 🛑
# @raycast.packageName Developer Tools
# @raycast.description Gracefully stops the local Win-Codex MCP server and Cloudflare tunnel processes.
# @raycast.author Deepak Bhattarai

$messages = @()

# 1. Stop Win-Codex Server
$conn = Get-NetTCPConnection -LocalPort 6000 -ErrorAction SilentlyContinue | Select-Object -First 1
if ($conn) {
    $proc = Get-Process -Id $conn.OwningProcess -ErrorAction SilentlyContinue
    if ($proc) {
        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        $messages += "Codex: stopped (PID $($proc.Id))"
    } else {
        $messages += "Codex: process stopped"
    }
} else {
    $messages += "Codex: not running"
}

# 2. Stop Cloudflare Tunnel
$cfProcs = Get-Process -Name "cloudflared" -ErrorAction SilentlyContinue
if ($cfProcs) {
    foreach ($p in $cfProcs) {
        Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
    }
    $messages += "Cloudflare: stopped"
} else {
    $messages += "Cloudflare: not running"
}

Write-Output ($messages -join " | ")
