#Requires AutoHotkey v2.0

; ==============================================================================
; TrayToast: 화면 우측 하단에 표시하는 자체 알림 상자
; - Windows 기본 알림(TrayTip)은 버튼을 넣거나 내용을 바꿀 수 없어 직접 그린다.
; - 알림을 클릭하면 같은 상자가 [열기] / [종료] 버튼 상자로 전환된다.
; - 마우스가 올라가 있는 동안은 자동으로 닫히지 않는다.
; ==============================================================================
class TrayToast {
    static Gui := 0
    static Expanded := false
    static Controls := Map()
    static OnOpen := 0
    static OnExit := 0
    static Fade := 0
    static AutoCloseFn := 0
    static FadeFn := 0
    static MessageHooked := false

    ; title: 굵은 제목, message: 안내 문구(클릭 시 버튼으로 바뀜)
    static Show(title, message, onOpen, onExit) {
        this.Close()
        this.OnOpen := onOpen
        this.OnExit := onExit
        this.Expanded := false

        if !this.AutoCloseFn {
            this.AutoCloseFn := ObjBindMethod(this, "AutoClose")
            this.FadeFn := ObjBindMethod(this, "FadeStep")
        }
        if !this.MessageHooked {
            ; 상자 빈 공간(배경) 클릭도 전환으로 처리한다.
            OnMessage(0x201, ObjBindMethod(this, "OnBackgroundClick"))
            this.MessageHooked := true
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
        g.SetFont("s9 norm c" sub)
        c["message"] := g.AddText("xm y+6 w304 BackgroundTrans", message)

        ; 전환 후 버튼: 안내 문구 자리에 표시한다.
        g.SetFont("s10 norm c" accentFg)
        c["open"] := g.AddText("xm yp w148 h34 Center 0x200 Hidden Background" accentBg, "열기")
        g.SetFont("s10 norm c" fg)
        c["exit"] := g.AddText("x+8 yp w148 h34 Center 0x200 Hidden Background" secondaryBg, "종료")

        for name in ["app", "title", "message"]
            c[name].OnEvent("Click", (*) => this.Expand())
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
        SetTimer(this.AutoCloseFn, -6000)
    }

    ; 알림 상자를 열기/종료 버튼 상자로 전환한다.
    static Expand() {
        if !this.Gui || this.Expanded
            return
        this.Expanded := true
        c := this.Controls
        c["message"].Visible := false
        c["open"].Visible := true
        c["exit"].Visible := true
        this.Gui.Show("NoActivate AutoSize")
        this.Place()
        SetTimer(this.AutoCloseFn, -10000)
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
        this.Expanded := false
    }

    static AutoClose() {
        if !this.Gui
            return
        ; 마우스가 상자 위에 있으면 닫지 않고 다시 대기한다.
        try {
            MouseGetPos(, , &hoverHwnd)
            if (hoverHwnd = this.Gui.Hwnd) {
                SetTimer(this.AutoCloseFn, -2000)
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

    static OnBackgroundClick(wParam, lParam, msg, hwnd) {
        if this.Gui && hwnd = this.Gui.Hwnd && !this.Expanded {
            this.Expand()
            return 0
        }
    }

    ; 작업 표시줄 위, 화면 우측 하단에 배치한다(하단 기준 고정).
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
