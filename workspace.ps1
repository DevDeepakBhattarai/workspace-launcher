Import-Module VirtualDesktop

# ---------------- APPS ----------------
# One app per virtual desktop, in order.

$apps = @(
    @{
        Key           = "Chrome"
        File          = "chrome.exe"
        Args          = @()
        WindowProcess = "chrome"
        TitleRegex    = "Chrome"
    },
    @{
        Key           = "Codex"
        File          = "explorer.exe"
        Args          = @("shell:AppsFolder\OpenAI.Codex_2p2nqsd0c76g0!App")
        WindowProcess = "Codex"
        TitleRegex    = "Codex"
    },
    @{
        Key           = "Zen"
        File          = "explorer.exe"
        Args          = @("shell:AppsFolder\F0DC299D809B9700")
        WindowProcess = "zen"
        TitleRegex    = "Zen"
    },
    @{
        Key           = "Antigravity"
        File          = "explorer.exe"
        Args          = @("shell:AppsFolder\Google.Antigravity")
        WindowProcess = "Antigravity"
        TitleRegex    = "Antigravity"
    },
    @{
        Key           = "Discord"
        File          = "Discord.exe"
        Args          = @()
        WindowProcess = "Discord"
        TitleRegex    = "Discord"
    },
    @{
        Key           = "WhatsApp"
        File          = "WhatsApp.exe"
        Args          = @()
        WindowProcess = "WhatsApp"
        TitleRegex    = "WhatsApp"
    }
)

# ---------------- WINDOWS API ----------------

Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public class Win32Window {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

    [DllImport("user32.dll")]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@

function Get-TopLevelWindows {
    $windows = New-Object System.Collections.Generic.List[object]

    $callback = [Win32Window+EnumWindowsProc]{
        param([IntPtr]$hWnd, [IntPtr]$lParam)

        if ([Win32Window]::IsWindowVisible($hWnd)) {
            $length = [Win32Window]::GetWindowTextLength($hWnd)
            if ($length -gt 0) {
                $builder = New-Object System.Text.StringBuilder ($length + 1)
                [void][Win32Window]::GetWindowText($hWnd, $builder, $builder.Capacity)

                [uint32]$procId = 0
                [void][Win32Window]::GetWindowThreadProcessId($hWnd, [ref]$procId)

                try {
                    $process = Get-Process -Id $pid -ErrorAction Stop
                    $windows.Add([pscustomobject]@{
                        Hwnd        = $hWnd
                        HwndInt     = $hWnd.ToInt64()
                        ProcessName = $process.ProcessName
                        Title       = $builder.ToString()
                    })
                } catch {}
            }
        }
        return $true
    }

    [void][Win32Window]::EnumWindows($callback, [IntPtr]::Zero)
    return $windows
}

function Wait-Window {
    param([object]$App, [hashtable]$Claimed, [int]$Timeout = 45)

    $deadline = (Get-Date).AddSeconds($Timeout)

    while ((Get-Date) -lt $deadline) {
        $match = Get-TopLevelWindows |
            Where-Object { $_.ProcessName -eq $App.WindowProcess -and -not $Claimed.ContainsKey($_.HwndInt) } |
            Where-Object { $_.Title -match $App.TitleRegex } |
            Select-Object -First 1

        if ($match) { return $match }
        Start-Sleep -Milliseconds 500
    }

    return $null
}

# ---------------- SETUP DESKTOPS ----------------

while ((Get-DesktopCount) -lt $apps.Count) {
    New-Desktop | Out-Null
    Start-Sleep -Milliseconds 250
}

# ---------------- LAUNCH ALL ----------------

foreach ($app in $apps) {
    Write-Host "Launching $($app.Key)..."
    try {
        Start-Process -FilePath $app.File -ArgumentList $app.Args
    } catch {
        Write-Warning "Failed to launch $($app.Key): $($_.Exception.Message)"
    }
}

# ---------------- MOVE TO DESKTOPS ----------------

$claimed = @{}

for ($i = 0; $i -lt $apps.Count; $i++) {
    $app = $apps[$i]
    Write-Host "Waiting for $($app.Key)..."

    $window = Wait-Window -App $app -Claimed $claimed

    if ($window) {
        Move-Window -Desktop (Get-Desktop $i) -Hwnd $window.Hwnd | Out-Null
        $claimed[$window.HwndInt] = $true
        Write-Host "$($app.Key) -> Desktop $($i + 1)"
    } else {
        Write-Warning "Could not find window for $($app.Key)"
    }
}

# Switch to Desktop 1
Get-Desktop 0 | Switch-Desktop -NoAnimation