; ============================================
; AUTO-ALLOW — klika Allow/Yes tylko gdy VS Code jest aktywne
; Ctrl+Shift+F12 = włącz/wyłącz
; ============================================

#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

INTERVAL := 250
BUTTONS := ["Allow", "Yes", "Zezwól", "Tak", "Allow All", "Zezwól na wszystko", "Trust", "Ufaj"]

TraySetIcon("shell32.dll", 294)
A_IconTip := "Auto-Allow aktywny"

SetTimer(AutoClick, INTERVAL)

AutoClick() {
    ; Tylko gdy VS Code jest aktywnym oknem
    try {
        if !WinActive("ahk_exe Code.exe")
            return
    } catch {
        return
    }

    hwnd := WinGetID("A")

    ; Metoda 1: szukaj przycisków po tytule
    for btn in BUTTONS {
        try {
            ControlClick(btn, hwnd)
            ToolTip("Auto-klik: " btn)
            SetTimer(() => ToolTip(), -1200)
            return
        }
    }

    ; Metoda 2: skanuj kontrolki po tekście
    try {
        for ctrl in WinGetControls(hwnd) {
            try {
                txt := ControlGetText(ctrl, hwnd)
                for btn in BUTTONS {
                    if InStr(txt, btn) {
                        ControlClick(ctrl, hwnd)
                        ToolTip("Auto-klik: " txt)
                        SetTimer(() => ToolTip(), -1200)
                        return
                    }
                }
            }
        }
    }
}

^+F12:: {
    static on := true
    on := !on
    SetTimer(AutoClick, on ? INTERVAL : 0)
    ToolTip("Auto-Allow: " (on ? "ON" : "OFF"))
    SetTimer(() => ToolTip(), -1500)
}
