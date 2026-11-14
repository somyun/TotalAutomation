#Requires AutoHotkey v2.0

; ==============================================================================
; TrayToast: 화면 우측 하단에 표시하는 자체 알림 상자
; - Windows 기본 알림(TrayTip)은 버튼을 넣을 수 없어 직접 그린다.
; - 처음부터 [열기] / [종료] 버튼이 있는 상자로 표시한다.
; - 아무 조작이 없으면 3초 뒤 닫히며, 마우스가 올라가 있는 동안은 닫지 않는다.
; ==============================================================================
class TrayToast {
    static Gui := 0
    static Controls := Map()
    static OnOpen := 0
    static OnExit := 0
    static Fade := 0
    static AutoCloseFn := 0
    static FadeFn := 0
    static DisplayMs := 3000

    static Show(title, onOpen, onExit) {
        this.Close()
        this.OnOpen := onOpen
        this.OnExit := onExit

        if !this.AutoCloseFn {
            this.AutoCloseFn := ObjBindMethod(this, "AutoClose")
            this.FadeFn := ObjBindMethod(this, "FadeStep")
        }

        dark := this.IsDarkTheme()
        bg := dark ? "2B2B2B" : "F9F9F9"
        fg := dark ? "FFFFFF" : "1B1B1B"
        sub := dark ? "C5C5C5" : "5C5C5C"
        secondaryBg := dark ? "3C3C3C" : "E5E5E5"
        accentBg := dark ? "4CC2FF" : "005FB8"
        accentFg := dark ? "000000" : "FFFFFF"

        g := Gui("+AlwaysOnTop -Caption +ToolWindow +Border", "통합자동화 알림")
        g.BackColor := bg
        g.MarginX := 16
        g.MarginY := 12

        c := Map()
        g.SetFont("s9 norm c" sub, "Malgun Gothic")
        c["app"] := g.AddText("xm ym w270 BackgroundTrans", "통합자동화")
        c["close"] := g.AddText("x+0 yp w34 Right BackgroundTrans", "✕")
        g.SetFont("s11 bold c" fg)
        c["title"] := g.AddText("xm y+8 w304 BackgroundTrans", title)
        g.SetFont("s10 norm c" accentFg)
        c["open"] := g.AddText("xm y+10 w148 h34 Center 0x200 Background" accentBg, "열기")
        g.SetFont("s10 norm c" fg)
        c["exit"] := g.AddText("x+8 yp w148 h34 Center 0x200 Background" secondaryBg, "종료")

        c["close"].OnEvent("Click", (*) => this.Close())
        c["open"].OnEvent("Click", (*) => this.Invoke("open"))
        c["exit"].OnEvent("Click", (*) => this.Invoke("exit"))

        this.Gui := g
        this.Controls := c

        g.Show("Hide AutoSize")
        this.ApplyRoundCorners(g.Hwnd)
        this.Place()
        WinSetTransparent(0, g.Hwnd)
        g.Show("NoActivate")
        this.Fade := 0
        SetTimer(this.FadeFn, 15)
        SetTimer(this.AutoCloseFn, -this.DisplayMs)
    }

    static Invoke(action) {
        callback := action = "open" ? this.OnOpen : this.OnExit
        this.Close()
        if IsObject(callback)
            callback.Call()
    }

    static Close() {
        if this.AutoCloseFn {
            SetTimer(this.AutoCloseFn, 0)
            SetTimer(this.FadeFn, 0)
        }
        if this.Gui {
            try this.Gui.Destroy()
            this.Gui := 0
        }
        this.Controls := Map()
    }

    static AutoClose() {
        if !this.Gui
            return
        ; 마우스가 상자 위에 있으면 닫지 않고 1초 뒤 다시 확인한다.
        try {
            MouseGetPos(, , &hoverHwnd)
            if (hoverHwnd = this.Gui.Hwnd) {
                SetTimer(this.AutoCloseFn, -1000)
                return
            }
        }
        this.Close()
    }

    static FadeStep() {
        if !this.Gui {
            SetTimer(this.FadeFn, 0)
            return
        }
        this.Fade := Min(this.Fade + 40, 255)
        try WinSetTransparent(this.Fade = 255 ? "Off" : this.Fade, this.Gui.Hwnd)
        if (this.Fade >= 255)
            SetTimer(this.FadeFn, 0)
    }

    ; 작업 표시줄 위, 화면 우측 하단에 배치한다.
    static Place() {
        if !this.Gui
            return
        WinGetPos(, , &w, &h, this.Gui.Hwnd)
        MonitorGetWorkArea(MonitorGetPrimary(), &left, &top, &right, &bottom)
        gap := Round(12 * A_ScreenDPI / 96)
        WinMove(right - w - gap, bottom - h - gap, , , this.Gui.Hwnd)
    }

    static ApplyRoundCorners(hwnd) {
        ; Windows 11: DWMWA_WINDOW_CORNER_PREFERENCE(33) = DWMWCP_ROUND(2). 이전 버전은 무시된다.
        try {
            pref := Buffer(4, 0)
            NumPut("Int", 2, pref)
            DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", hwnd, "UInt", 33, "Ptr", pref, "UInt", 4)
        }
    }

    static IsDarkTheme() {
        try return RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
            "SystemUsesLightTheme") = 0
        return false
    }
}
