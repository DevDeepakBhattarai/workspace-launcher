#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

; DJI Mic Mini receiver: VID 2CA3 / PID 4011.
; Its link/shutter button arrives as a HID Consumer Control Volume Up event.
; We correlate the raw HID source with the Volume_Up hotkey so only the DJI
; single press becomes Win+Space; a double press toggles ChatGPT microphone
; through Super App. Normal keyboard/headset volume keys keep working.

global g_LastDjiConsumerInput := 0
global g_PendingDjiPress := false
global g_DeviceCache := Map()
global g_RawInputGui := Gui("+ToolWindow -Caption")
g_RawInputGui.Show("Hide")

RegisterDjiConsumerRawInput()

$Volume_Up::HandleVolumeUp()

RegisterDjiConsumerRawInput() {
    global g_RawInputGui

    ; RAWINPUTDEVICE for Consumer Controls: Usage Page 0x0C, Usage 0x01.
    ridSize := (A_PtrSize = 8) ? 16 : 12
    rid := Buffer(ridSize, 0)
    NumPut("UShort", 0x0C, rid, 0)
    NumPut("UShort", 0x01, rid, 2)
    NumPut("UInt", 0x00000100, rid, 4) ; RIDEV_INPUTSINK
    NumPut("Ptr", g_RawInputGui.Hwnd, rid, 8)

    if !DllCall("User32\RegisterRawInputDevices", "Ptr", rid, "UInt", 1, "UInt", ridSize) {
        MsgBox "Could not register for HID Consumer Control raw input.", "DJI Mic Voice Hotkey", "Iconx"
        ExitApp
    }

    OnMessage(0x00FF, OnRawInput) ; WM_INPUT
}

OnRawInput(wParam, lParam, msg, hwnd) {
    global g_LastDjiConsumerInput

    static RID_INPUT := 0x10000003
    static RIM_TYPEHID := 2
    headerSize := 8 + (2 * A_PtrSize)
    dataSize := 0

    if DllCall("User32\GetRawInputData", "Ptr", lParam, "UInt", RID_INPUT,
        "Ptr", 0, "UIntP", &dataSize, "UInt", headerSize) = 0xFFFFFFFF
        return

    if dataSize < headerSize
        return

    raw := Buffer(dataSize, 0)
    if DllCall("User32\GetRawInputData", "Ptr", lParam, "UInt", RID_INPUT,
        "Ptr", raw, "UIntP", &dataSize, "UInt", headerSize) = 0xFFFFFFFF
        return

    if NumGet(raw, 0, "UInt") != RIM_TYPEHID
        return

    hDevice := NumGet(raw, 8, "Ptr")
    if IsDjiMicMiniDevice(hDevice)
        g_LastDjiConsumerInput := A_TickCount
}

IsDjiMicMiniDevice(hDevice) {
    global g_DeviceCache
    static RIDI_DEVICENAME := 0x20000007

    if g_DeviceCache.Has(hDevice)
        return g_DeviceCache[hDevice]

    chars := 0
    result := DllCall("User32\GetRawInputDeviceInfoW", "Ptr", hDevice, "UInt", RIDI_DEVICENAME,
        "Ptr", 0, "UIntP", &chars)
    if result = 0xFFFFFFFF || chars = 0 {
        g_DeviceCache[hDevice] := false
        return false
    }

    nameBuf := Buffer((chars + 1) * 2, 0)
    result := DllCall("User32\GetRawInputDeviceInfoW", "Ptr", hDevice, "UInt", RIDI_DEVICENAME,
        "Ptr", nameBuf, "UIntP", &chars)
    if result = 0xFFFFFFFF {
        g_DeviceCache[hDevice] := false
        return false
    }

    deviceName := StrUpper(StrGet(nameBuf, chars, "UTF-16"))
    isDji := InStr(deviceName, "VID_2CA3&PID_4011") != 0
    g_DeviceCache[hDevice] := isDji
    return isDji
}

HandleVolumeUp() {
    global g_LastDjiConsumerInput, g_PendingDjiPress

    ; Give WM_INPUT time to identify the source device.
    Sleep 35
    if (A_TickCount - g_LastDjiConsumerInput <= 180) {
        if g_PendingDjiPress {
            SetTimer SendDjiTranscription, 0
            g_PendingDjiPress := false
            SendEvent "^!+m"
        } else {
            g_PendingDjiPress := true
            SetTimer SendDjiTranscription, -350
        }
        ; Ignore key repeat until this physical press is released.
        KeyWait "Volume_Up"
        return
    }

    SendEvent "{Volume_Up}"
}

SendDjiTranscription() {
    global g_PendingDjiPress
    if !g_PendingDjiPress
        return
    g_PendingDjiPress := false
    SendEvent "#{Space}"
}
