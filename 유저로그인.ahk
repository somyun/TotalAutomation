class UserLogin {
    static LoginUser() {
        return this.selectAndAuthUser()
    }

    static selectAndAuthUser() {
        loop {
            guiObj := UserSelectionGUI()
            selectedUser := guiObj.WaitForSubmit()

            if (!selectedUser || !selectedUser.Has("id")) {
                return false
            }

            ; 프로필 확인
            profiles := ConfigManager.GetProfiles()
            targetProfile := ""
            for p in profiles {
                if p["id"] == selectedUser["id"] {
                    targetProfile := p
                    break
                }
            }

            if !targetProfile {
                MsgBox "해당 유저 프로필을 찾을 수 없습니다.", "오류", "Iconx"
                continue
            }

            ; 2차 비밀번호 확인
            ; (기존 단축일지처럼 2차 비번으로 본인 인증)
            pwBox := InputBox("2차 비밀번호를 입력하세요", "인증", "Password w250")
            if (pwBox.Result != "OK") {
                continue
            }

            savedPW2 := targetProfile.Has("pw2") ? targetProfile["pw2"] : ""

            ; 2차 비밀번호가 설정되지 않은 경우 신규 유저로 간주하여 통과시킬 수도 있으나,
            ; 등록 시 필수로 입력받으므로 불일치면 실패 처리
            if (savedPW2 != "" && pwBox.Value != savedPW2) {
                MsgBox "2차 비밀번호가 일치하지 않습니다.", "오류", "Icon!"
                continue
            }

            ; 로그인 성공
            ConfigManager.Set("appSettings.lastUser", targetProfile["id"]) ; ID로 저장 (이름 대신)
            ConfigManager.CurrentUser := targetProfile
            return true
        }
    }
}

class UserSelectionGUI {
    controls := Map()
    result := Map()
    submitted := false

    __New() {
        this.controls["main"] := gui("", "사용자 선택")
        this.controls["main"].SetFont("S10", "맑은 고딕")
        this.controls["main"].Opt("-MinimizeBox -MaximizeBox")

        this.controls["main"].AddText("Section", "등록된 유저:")
        this.controls["userList"] := this.controls["main"].AddListBox("xs w200 h150")
        this.controls["userList"].OnEvent("DoubleClick", (*) => this.onSelectUser())

        ; 버튼 그룹
        this.controls["btn_ok"] := this.controls["main"].AddButton("xs w200 h30", "로그인")
        this.controls["btn_ok"].OnEvent("Click", (*) => this.onSelectUser())

        this.controls["btn_add"] := this.controls["main"].AddButton("xs w95 h30", "유저 추가")
        this.controls["btn_add"].OnEvent("Click", (*) => this.onAddUser())

        this.controls["btn_del"] := this.controls["main"].AddButton("x+10 w95 h30", "삭제")
        this.controls["btn_del"].OnEvent("Click", (*) => this.onDeleteUser())

        this.loadUserList()
        this.controls["main"].OnEvent("Close", (*) => this.onCancel())
        this.controls["main"].Show("Center")
    }

    loadUserList() {
        profiles := ConfigManager.GetProfiles()
        this.controls["userList"].Delete()
        lastUserId := ConfigManager.Get("appSettings.lastUser")

        selectIndex := 0
        for i, p in profiles {
            name := p.Has("name") ? p["name"] : p["id"]
            this.controls["userList"].Add([name " (" p["id"] ")"])
            if (p["id"] == lastUserId)
                selectIndex := i
        }

        if (selectIndex > 0)
            this.controls["userList"].Choose(selectIndex)
    }

    onSelectUser() {
        idx := this.controls["userList"].Value
        if (idx == 0) {
            MsgBox "유저를 선택해주세요."
            return
        }

        profiles := ConfigManager.GetProfiles()
        if (idx > profiles.Length)
            return

        this.result := profiles[idx] ; 1-based index
        this.submitted := true
        this.controls["main"].Destroy()
    }

    onAddUser() {
        this.controls["main"].Opt("+Disabled") ; 메인창 비활성
        regGui := UserInputGUI()
        newUserResult := regGui.WaitForSubmit()
        this.controls["main"].Opt("-Disabled") ; 메인창 활성
        this.controls["main"].Show() ; 다시 포커스

        if (newUserResult && newUserResult.Count > 0) {
            newID := newUserResult["id"]

            if ConfigManager.Config["users"].Has(newID) {
                MsgBox "이미 존재하는 사번입니다: " newID, "오류", "Iconx"
                return
            }

            ; Users 객체에 추가 (newUserResult가 이미 구조화된 Map임)
            ConfigManager.Config["users"][newID] := newUserResult
            ConfigManager.Save()
            this.loadUserList()
        }
    }

    ; 유저 삭제
    onDeleteUser() {
        idx := this.controls["userList"].Value
        if (idx == 0) {
            return
        }

        profiles := ConfigManager.GetProfiles()
        targetProfile := profiles[idx]
        targetID := targetProfile["id"]

        if MsgBox("'" targetProfile["name"] "' 님의 정보를 삭제하시겠습니까?", "삭제 확인", "YesNo Icon?") == "Yes" {
            ; 삭제 시 2차 비밀번호 확인
            pwBox := InputBox("삭제하려면 2차 비밀번호를 입력하세요", "인증", "Password w250")
            if (pwBox.Result != "OK") {
                return
            }

            savedPW2 := targetProfile.Has("pw2") ? targetProfile["pw2"] : ""
            if (savedPW2 != "" && pwBox.Value != savedPW2) {
                MsgBox "2차 비밀번호가 일치하지 않습니다.", "오류", "Icon!"
                return
            }

            ; ConfigManager에서 유저 삭제 (users 객체에서 해당 키 삭제)
            if ConfigManager.Config.Has("users") && ConfigManager.Config["users"].Has(targetID) {
                ConfigManager.Config["users"].Delete(targetID)
                ConfigManager.Save()
            }
            this.loadUserList()
        }
    }

    onCancel() {
        this.result := Map()
        this.submitted := false
        this.controls["main"].Destroy()
    }

    WaitForSubmit() {
        while WinExist("사용자 선택")
            Sleep 100
        return this.result
    }
}

class UserInputGUI {
    controls := Map()
    result := Map()
    submitted := false

    __New() {
        this.controls["main"] := Gui("+Owner", "유저 정보 등록")
        this.controls["main"].SetFont("S10", "맑은 고딕")
        this.controls["main"].Opt("-MinimizeBox -MaximizeBox")

        mainGui := this.controls["main"]

        mainGui.AddText("w120 Section", "이름 *")
        mainGui.AddText("w120", "사번 *")
        mainGui.AddText("w120", "통합 PW *")
        mainGui.AddText("w120", "통합 PW 확인 *")
        mainGui.AddText("w120", "2차 PW *")
        mainGui.AddText("w120", "2차 PW 확인 *")
        mainGui.AddText("w120", "SAP PW")
        mainGui.AddText("w120", "SAP PW 확인")

        this.controls["name"] := mainGui.AddEdit("ys-3 Section w120")
        this.controls["webID"] := mainGui.AddEdit("w120 Number Limit6")

        ; 근무조 자동계산 로직은 복잡하니 일단 선택으로
        this.controls["team"] := mainGui.AddDropDownList("x+10 yp w80 Choose1", ["A조", "B조", "C조", "D조", "일근"])

        this.controls["webPW1"] := mainGui.AddEdit("xs Password w210")
        this.controls["webPW2"] := mainGui.AddEdit("Password w210")
        this.controls["pw2_1"] := mainGui.AddEdit("Password Number Limit6 w210")
        this.controls["pw2_2"] := mainGui.AddEdit("Password Number Limit6 w210")
        this.controls["sapPW1"] := mainGui.AddEdit("Password w210")
        this.controls["sapPW2"] := mainGui.AddEdit("Password w210")

        ; 체크 표시용
        this.setupPwCheck("webPW1", "webPW2")
        this.setupPwCheck("pw2_1", "pw2_2")
        this.setupPwCheck("sapPW1", "sapPW2")

        btnSave := mainGui.AddButton("xs w210 h35", "저장")
        btnSave.OnEvent("Click", (*) => this.onSave())

        mainGui.OnEvent("Close", (*) => this.onCancel())
        mainGui.Show("Center")
    }

    setupPwCheck(id1, id2) {
        this.controls[id1].OnEvent("Change", (*) => this.checkMatch(id1, id2))
        this.controls[id2].OnEvent("Change", (*) => this.checkMatch(id1, id2))
    }

    checkMatch(id1, id2) {
        val1 := this.controls[id1].Value
        val2 := this.controls[id2].Value

        if (val1 != "" && val2 != "" && val1 == val2)
            this.controls[id2].Opt("+cGreen")
        else
            this.controls[id2].Opt("+cBlack")
    }

    onSave() {
        name := Trim(this.controls["name"].Value)
        id := Trim(this.controls["webID"].Value)
        team := this.controls["team"].Text
        wp1 := this.controls["webPW1"].Value
        wp2 := this.controls["webPW2"].Value
        p2_1 := this.controls["pw2_1"].Value
        p2_2 := this.controls["pw2_2"].Value
        sp1 := this.controls["sapPW1"].Value
        sp2 := this.controls["sapPW2"].Value

        if (name == "" || id == "" || wp1 == "" || wp2 == "" || p2_1 == "" || p2_2 == "") {
            MsgBox "필수 항목(*)을 모두 입력해주세요.", "알림"
            return
        }

        if (wp1 != wp2) {
            MsgBox "통합 비밀번호가 일치하지 않습니다.", "오류"
            return
        }
        if (p2_1 != p2_2) {
            MsgBox "2차 비밀번호가 일치하지 않습니다.", "오류"
            return
        }
        if (sp1 != "" && sp1 != sp2) {
            MsgBox "SAP 비밀번호가 일치하지 않습니다.", "오류"
            return
        }

        ; 저장할 데이터 Map 생성 (ConfigManager 구조에 맞게)
        this.result := Map(
            "id", id,
            "profile", Map(
                "id", id,
                "name", name,
                "webPW", wp1,
                "pw2", p2_1,
                "sapPW", sp1,
                "team", team,
                "department", "호포전기분소"
            ),
            "hotkeys", [],
            "presets", Map()
        )

        this.submitted := true
        this.controls["main"].Destroy()
    }

    onCancel() {
        this.result := Map()
        this.submitted := false
        this.controls["main"].Destroy()
    }

    WaitForSubmit() {
        while WinExist("유저 정보 등록")
            Sleep 100
        return this.result
    }
}

; ==============================================================================
; 웹 자동 로그인 관리 클래스
; ==============================================================================
class WebAutoLogin {
    static PortalURL := "https://btcep.humetro.busan.kr/portal"
    static ERP_PortalURL := "https://niw.humetro.busan.kr/erpep.jsp"
    static Worklog_List :=
        "http://ep.humetro.busan.kr/irj/portal?NavigationTarget=ROLES%3A%2F%2Fportal_content%2Fhumetro%2Frole%2Fmaintenance%2Frole.09%2Fworkset.07%2Fworkset.01%2Fworkset.03&sapDocumentRenderingMode=EmulateIE8"

    ; ==============================================================================
    ; [메서드] EnsureReady
    ; 설명: 작업 유형에 따른 브라우저 상태를 준비합니다.
    ; ==============================================================================
    static EnsureReady(taskType) {
        user := ConfigManager.CurrentUser
        if (!user.Has("id")) {
            MsgBox("로그인된 사용자가 없습니다. 먼저 로컬 로그인을 수행해주세요.", "오류", "Iconx")
            return false
        }

        ; 1. 쿠키 데이터 확인
        if (HeadlessAutomation.CookieStorage.Count == 0) {
            MsgBox("로그인 정보(쿠키)가 없습니다. 백그라운드 로그인이 완료되었는지 확인해주세요.", "오류", "Iconx")
            return false
        }

        if (taskType == "SessionCheck") {
            ; '선로출입관리', '승인정보불러오기' 등 XPLATFORM 구동용 세션 연동
            return this.LaunchXPlatformSession(user)
        }
        else if (taskType == "WorkLog_Create") {
            try {
                return this.LaunchLogSession(user, "reg", "general")
            } catch as e {
                MsgBox("업무일지 생성 화면 이동 중 오류: " e.Message, "오류", "Iconx")
                return false
            }
        }
        else if (taskType == "WorkLog_View") {
            try {
                return this.LaunchLogSession(user, "mod", "general")
            } catch as e {
                MsgBox("업무일지 조회 화면 이동 중 오류: " e.Message, "오류", "Iconx")
                return false
            }
        }

        return true
    }

    ; ==============================================================================
    ; [메서드] LaunchLogSession
    ; 설명: 쿠키 주입 후 Edge 브라우저를 실행하여 일지 작성/조회 창을 엽니다.
    ; ==============================================================================
    static LaunchLogSession(user, mode := "mod", browserMode := "general") {
        ; url 준비
        targetUrl := ""
        iljino := ""

        if (mode == "mod") {
            headless := HeadlessAutomation(true)
            if (!user.Has("arbpl") || user["arbpl"] == "") {
                MsgBox("부서코드(arbpl)를 불러올 수 없습니다.`n(사번: " user["id"] ")")
                return false
            }
            iljino := headless.GetTodayWorkLogNumber(user["id"], user["arbpl"])

            if (iljino == "") {
                MsgBox("형식에 맞는 오늘자 일지 번호를 찾을 수 없습니다.`n(부서코드: " user["arbpl"] ")")
                return false
            }

            targetUrl :=
                "http://ep.humetro.busan.kr/irj/servlet/prt/portal/prtroot/kr.busan.humetro.cbo.erp.work_log.WorkLogReg"
                . "?I_MODE=MOD&V_ILJINO=" iljino "&V_SABUN=" user["id"]
        } else if (mode == "reg") {
            today := FormatTime(DateAdd(A_Now, -9, "Hours"), "yyyy-MM-dd")
            targetUrl :=
                "http://ep.humetro.busan.kr/irj/servlet/prt/portal/prtroot/kr.busan.humetro.cbo.erp.work_log.WorkLogReg"
                . "?I_MODE=REG&V_SABUN=" user["id"]
                . "&V_ARBPL01=" user["arbpl"]
                . "&I_ARWRK=5010"
                . "&I_GIJUNDF=" today
                . "&I_GIJUNDT=" today
        } else {
            return false
        }

        ; 브라우저 인자 조립
        edgePath := "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
        profile := A_Temp "\edge_cookie_profile" A_TickCount

        args := ' --remote-debugging-port=9222 --user-data-dir="' profile '"'
            . ' --no-first-run --no-default-browser-check --disable-default-apps'

        if (browserMode == "app") {
            args .= ' "data:text/html;charset=utf-8,Loading..."'
        } else {
            args .= ' about:blank'
        }

        try {
            Run(Format('"{1}" {2}', edgePath, args), , "max", &edgePid)
        } catch as e {
            MsgBox("브라우저 실행 실패: " e.Message, "오류", "Iconx")
            return false
        }

        try {
            ; Chrome 인스턴스 연결
            chromeInst := Chrome([], , , 9222)
            pages := chromeInst.GetPageList()
            foundData := ""

            for p in pages {
                if (p.Has("type") && p["type"] == "page") {
                    if (p.Has("webSocketDebuggerUrl")) {
                        foundData := p
                        break
                    }
                }
            }

            if (foundData) {
                wsUrl := StrReplace(foundData["webSocketDebuggerUrl"], "localhost", "127.0.0.1")
                page := Chrome.Page(wsUrl)

                cookieParams := HeadlessAutomation.GetCookieParamsForCDP()
                page.Call("Network.enable")
                page.Call("Network.setCookies", Map("cookies", cookieParams))

                if (browserMode == "app") {
                    js := Format(
                        "window.open('{1}', '_blank', 'width=1024,height=760,menubar=no,toolbar=no,location=yes,status=yes,scrollbars=yes,resizable=yes');",
                        targetUrl)
                    page.Evaluate(js)
                    Sleep(500)
                    page.Call("Page.close")

                } else {
                    DirectURL :=
                        "http://ep.humetro.busan.kr/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fhumetro!2frole!2fmaintenance!2frole.09!2fworkset.07!2fworkset.01!2fworkset.03!2fiview.02?sapDocumentRenderingMode=EmulateIE8"
                    page.Call("Page.navigate", Map("url", DirectURL))

                    modeCode := (mode == "mod") ? "MOD" : "REG"

                    js := "let count = 0; "
                        . "let inter = setInterval(() => { "
                        . "    if (typeof $ !== 'undefined' && typeof fn_detail_open !== 'undefined') { "
                        . "        clearInterval(inter); "

                    ; 조회 모드일 때만 일지번호 세팅
                    if (mode == "mod" && iljino != "")
                        js .= "        $('#V_ILJINO').val('" iljino "'); "

                    js .= "        $('#I_MODE').val('" modeCode "'); "
                        .
                        "        fn_detail_open('WorkLogForm', 'reg', 'kr.busan.humetro.cbo.erp.work_log.WorkLogReg', 'width=1024,height=760,scrollbars=yes,resizable=yes,status=yes'); "
                        . "    } else if (count++ > 20) { "
                        . "        clearInterval(inter); "
                        . "    } "
                        . "}, 500);"

                    page.Evaluate(js)
                }
            }
        } catch as e {
            MsgBox("브라우저 제어(쿠키/이동) 오류: " e.Message, "오류", "Iconx")
            return false
        }

        ; 열린 팝업창(일지 상세창)의 UIA 객체를 획득하여 반환
        try {
            ; JS(fn_detail_open) 실행 후 창이 2개(메인, 팝업)가 될 때까지 최대 15초 대기
            loop 30 {
                Sleep 100
                hwnds := WinGetList("ahk_pid " edgePid)
                if (hwnds.Length >= 2) {
                    WinActivate("ahk_id " hwnds[1]) ; 팝업창 활성화 보장
                    return UIA_Browser("ahk_id " hwnds[1])
                }
            }
            MsgBox("일지 상세창(팝업) 호출 시간 초과", "알림", "Iconx")
            return false
        } catch as e {
            MsgBox("cUIA 연결 실패: " e.Message, "오류", "Iconx")
            return false
        }
    }

    ; ==============================================================================
    ; [메서드] LaunchXPlatformSession
    ; 설명: XPLATFORM 로컬 런처 (7936 포트)와 통신하여 쿠키를 주입하고 XPLATFORM을 구동합니다.
    ; ==============================================================================
    static LaunchXPlatformSession(user) {
        http := ComObject("WinHttp.WinHttpRequest.5.1")
        http.Option[3] := 0  ; EnableCookieExchange = False
        http.Option[4] := 13056 ; IgnoreCertErrors

        try {
            launcherUrl := "https://127.0.0.1:7936/launcher/xplatform/" A_TickCount

            ; 1. Create
            http.Open("POST", launcherUrl, false)
            http.SetRequestHeader("Content-Type", "text/plain;charset=UTF-8")
            http.Send('{"platform":"xplatform","action":"create"}')

            createResp := JSON.parse(http.ResponseText)
            if (createResp.Has("result") && createResp["result"] != "success")
                throw Error("런처 세션 생성 실패")

            launchId := createResp["id"]

            ; 2. setproperty
            setPropVal := Map(
                "key", "eGovXplatform",
                "xadl", "https://mis.humetro.busan.kr:443/FS/xui/XUI.xadl",
                "componentpath", "%USERAPPLOCAL%\XPLATFORM\921\component\",
                "cachedir", "%CACHE%",
                "splashimage", "https://mis.humetro.busan.kr:443/FS/xui//install/img/loading_img.gif",
                "loadingimage", "https://mis.humetro.busan.kr:443/FS/xui//install/img/loading_img.gif",
                "commthreadwaittime", 1,
                "version", "9.2.1",
                "commthreadcount", 3,
                "errorfile", "xplatform.xml",
                "globalvalue", "gv_selSystGubn=LA,gv_userBrowser=Edg,,gv_ssoUserId=" user["id"] ",gv_svcUrl=https://mis.humetro.busan.kr:443/FS/",
                "onlyone", false,
                "showiniterror", false,
                "usewininet", true,
                "enginesetupkey", "{AA890DB4-7357-4237-82BB-D0B931AAB420}",
                "splashmessage", "new lanucher test..."
            )
            setPropPayload := JSON.stringify(Map("platform", "xplatform", "action", "setproperty", "id", launchId,
                "value", setPropVal))

            http.Open("POST", launcherUrl, false)
            http.Send(setPropPayload)

            ; 3. addWebInfo (ep 도메인 쿠키)
            epCookieStr := HeadlessAutomation.CookieStorage.Has("ep") ? HeadlessAutomation.CookieStorage["ep"] : ""
            if (epCookieStr == "") {
                MsgBox("XPLATFORM 구동에 필요한 ep 쿠키 정보를 찾을 수 없습니다.", "에러", "Iconx")
                return false
            }

            addWebInfoPayload := JSON.stringify(Map(
                "platform", "xplatform",
                "action", "method",
                "id", launchId,
                "value", Map("addWebInfo", Map("param", [epCookieStr]))
            ))

            http.Open("POST", launcherUrl, false)
            http.Send(addWebInfoPayload)

            ; 4. launch
            launchPayload := JSON.stringify(Map(
                "platform", "xplatform",
                "action", "method",
                "id", launchId,
                "value", Map("launch", "ok")
            ))

            http.Open("POST", launcherUrl, false)
            http.Send(launchPayload)

            return true

        } catch as err {
            MsgBox("XPLATFORM 실행 중 오류 발생 (런처 데몬이 종료되었거나 포트가 다릅니다): " err.Message, "오류", "Iconx")
            return false
        }
    }
}
