#Requires AutoHotkey v2.0
#SingleInstance Force

; =====================================================
; VirtualDesktopAccessor.dll
; =====================================================

DllPath := A_ScriptDir "\VirtualDesktopAccessor.dll"

if !FileExist(DllPath) {
    MsgBox(
        "VirtualDesktopAccessor.dll not found!`n`nDownload from:`nhttps://github.com/Ciantic/VirtualDesktopAccessor/releases",
        "Error", "Icon!")
    ExitApp
}

DllCall("LoadLibrary", "Str", DllPath)

GetDesktopCount() => DllCall(DllPath "\GetDesktopCount", "Int")
GoToDesktopNumber(n) => DllCall(DllPath "\GoToDesktopNumber", "Int", n)
MoveWindowToDesktopNumber(hwnd, n)
    => DllCall(DllPath "\MoveWindowToDesktopNumber", "Ptr", hwnd, "Int", n)

; =====================================================
; WORKSPACE CONFIG (0-based desktops)
; =====================================================

Apps := [{ name: "Chrome", exe: "chrome.exe", path: "C:\Program Files\Google\Chrome\Application\chrome.exe", desktop: 0,
    fullscreen: true }, { name: "Cursor", exe: "Cursor.exe", path: "C:\Users\deepak_bhattarai\AppData\Local\Programs\cursor\Cursor.exe",
        desktop: 1, fullscreen: true }, { name: "Zen Browser", exe: "zen.exe", path: "C:\Program Files\Zen Browser\zen.exe",
            desktop: 2 }, { name: "Discord", exe: "Discord.exe", path: "C:\Users\deepak_bhattarai\AppData\Local\Discord\Update.exe",
                args: "--processStart Discord.exe", desktop: 4 }, { name: "Raindrop", exe: "Raindrop.io.exe", path: "shell:AppsFolder\19059Raindrop.io.Raindrop.io_hghhavmbrcx2t!App",
                    desktop: 3, tile: "right" }, { name: "TickTick", exe: "TickTick.exe", path: "C:\Program Files (x86)\TickTick\TickTick.exe",
                        desktop: 3, tile: "left" }
]

; =====================================================
; HELPERS
; =====================================================

EnsureDesktopsExist(count) {
    while GetDesktopCount() < count {
        Send "^#d"
        Sleep 120
    }
}

LaunchApp(app) {
    if ProcessExist(app.exe)
        return
    try {
        if app.HasProp("args")
            Run('"' app.path '" ' app.args)
        else
            Run('"' app.path '"')
    }
}

FindAppWindow(app) {
    ; Try winTitle first (for UWP/Store apps)
    if app.HasProp("winTitle") {
        hwnd := WinExist(app.winTitle)
        if hwnd
            return hwnd
    }
    ; Fall back to exe matching
    return WinExist("ahk_exe " app.exe)
}

; =====================================================
; SIMPLE WINDOW TILING (PRIMARY MONITOR)
; =====================================================

TileWindow(hwnd, side := "left") {
    ; Use primary monitor
    MonitorGetWorkArea(1, &L, &T, &R, &B)
    mw := R - L
    mh := B - T

    if (side = "left")
        WinMove L, T, mw // 2, mh, hwnd
    else
        WinMove L + mw // 2, T, mw // 2, mh, hwnd
}

; =====================================================
; CORE LOGIC
; =====================================================

LaunchAndOrganize() {
    EnsureDesktopsExist(5)

    ; Launch all apps first
    for app in Apps
        LaunchApp(app)

    ; Brief wait for apps to start
    Sleep 2000

    ; Move and arrange whatever windows exist
    for app in Apps {
        hwnd := FindAppWindow(app)
        if !hwnd
            continue

        MoveWindowToDesktopNumber(hwnd, app.desktop)

        if app.HasProp("tile")
            TileWindow(hwnd, app.tile)
        else if app.HasProp("fullscreen") && app.fullscreen
            WinMaximize hwnd
    }

    GoToDesktopNumber 0
}

OrganizeOnly() {
    EnsureDesktopsExist(5)

    for app in Apps {
        hwnd := FindAppWindow(app)
        if !hwnd
            continue

        MoveWindowToDesktopNumber(hwnd, app.desktop)

        if app.HasProp("tile")
            TileWindow(hwnd, app.tile)
        else if app.HasProp("fullscreen") && app.fullscreen
            WinMaximize hwnd
    }

    GoToDesktopNumber 0
}

MoveActiveToDesktop(n) {
    hwnd := WinGetID("A")
    if hwnd {
        MoveWindowToDesktopNumber(hwnd, n)
        GoToDesktopNumber n
    }
}

; =====================================================
; HOTKEYS
; =====================================================

#!o:: LaunchAndOrganize()
#!p:: OrganizeOnly()

; =====================================================
; CLOSE ALL WINDOWS (Across All Virtual Desktops)
; =====================================================

GetWindowDesktopNumber(hwnd) => DllCall(DllPath "\GetWindowDesktopNumber", "Ptr", hwnd, "Int")

CloseAllWindows() {
    global DllPath

    ; List of processes to exclude (system/shell processes)
    excludeProcesses := ["explorer.exe", "SystemSettings.exe", "ShellExperienceHost.exe",
        "SearchHost.exe", "StartMenuExperienceHost.exe", "TextInputHost.exe",
        "AutoHotkey64.exe", "AutoHotkey32.exe", "AutoHotkey.exe",
        "Taskmgr.exe", "ApplicationFrameHost.exe", "LockApp.exe",
        "WindowsTerminal.exe", "powershell.exe", "cmd.exe"]

    closedCount := 0
    desktopCount := GetDesktopCount()
    windowsToClose := []

    ; Use EnumWindows to get ALL windows including those on other desktops
    EnumWindowsProc := CallbackCreate(EnumWindowsCallback, "F", 2)

    ; Create a structure to pass data to callback
    global g_WindowsToClose := []
    global g_ExcludeProcesses := excludeProcesses
    global g_DesktopCount := desktopCount

    ; EnumWindows enumerates all top-level windows on all desktops
    DllCall("EnumWindows", "Ptr", EnumWindowsProc, "Ptr", 0)
    CallbackFree(EnumWindowsProc)

    ; Now close all collected windows
    for hwnd in g_WindowsToClose {
        try {
            WinClose(hwnd)
            closedCount++
            Sleep 30
        }
    }

    ; Cleanup globals
    g_WindowsToClose := []

    ; Show notification
    if (closedCount > 0)
        ToolTip("Closed " closedCount " window(s) across all desktops")
    else
        ToolTip("No windows to close")

    SetTimer () => ToolTip(), -2000
}

EnumWindowsCallback(hwnd, lParam) {
    global DllPath, g_WindowsToClose, g_ExcludeProcesses, g_DesktopCount

    try {
        ; Skip invisible windows
        if !DllCall("IsWindowVisible", "Ptr", hwnd)
            return true

        ; Get window desktop number
        windowDesktop := DllCall(DllPath "\GetWindowDesktopNumber", "Ptr", hwnd, "Int")

        ; Skip windows not on any valid desktop (-1 means system window)
        if (windowDesktop < 0 || windowDesktop >= g_DesktopCount)
            return true

        ; Get process name
        try {
            processName := WinGetProcessName(hwnd)
        } catch {
            return true
        }

        ; Skip excluded processes
        for excluded in g_ExcludeProcesses {
            if (StrLower(processName) = StrLower(excluded))
                return true
        }

        ; Skip windows without titles
        try {
            title := WinGetTitle(hwnd)
            if (title = "")
                return true
        } catch {
            return true
        }

        ; Check if it's a real top-level window (has WS_VISIBLE and no owner)
        style := DllCall("GetWindowLong", "Ptr", hwnd, "Int", -16, "Int")  ; GWL_STYLE
        if !(style & 0x10000000)  ; WS_VISIBLE
            return true

        ; Check if window has an owner (skip owned windows like tooltips)
        owner := DllCall("GetWindow", "Ptr", hwnd, "UInt", 4, "Ptr")  ; GW_OWNER
        if (owner != 0)
            return true

        ; Add to list of windows to close
        g_WindowsToClose.Push(hwnd)
    }

    return true  ; Continue enumeration
}

#!1:: MoveActiveToDesktop(0)
#!2:: MoveActiveToDesktop(1)
#!3:: MoveActiveToDesktop(2)
#!4:: MoveActiveToDesktop(3)
#!5:: MoveActiveToDesktop(4)
#!q:: CloseAllWindows()