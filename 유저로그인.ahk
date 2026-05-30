; ==============================================================================
; 웹 자동 로그인 관리 클래스
; ==============================================================================
class WebAutoLogin {
    static PortalURL := "https://btcep.humetro.busan.kr/portal"
    static ERP_PortalURL := "https://niw.humetro.busan.kr/erpep.jsp"
    static Worklog_List :=
        "http://ep.humetro.busan.kr/irj/portal?NavigationTarget=ROLES%3A%2F%2Fportal_content%2Fhumetro%2Frole%2Fmaintenance%2Frole.09%2Fworkset.07%2Fworkset.01%2Fworkset.03&sapDocumentRenderingMode=EmulateIE8"
    ;static Worklog_List_Direct :="http://ep.humetro.busan.kr/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fhumetro!2frole!2fmaintenance!2frole.09!2fworkset.07!2fworkset.01!2fworkset.03!2fiview.02?sapDocumentRenderingMode=EmulateIE8" ;헤더미포함

    ; ==============================================================================
    ; [메서드] EnsureReady
    ; 설명: 작업 유형에 따른 브라우저 상태를 준비합니다.
    ; ==============================================================================
    static EnsureReady(taskType) {
        user := ConfigManager.CurrentUser
        if (!user.Has("id")) {
            LogDebug("[오류] 로그인된 사용자가 없음 (EnsureReady)")
            MsgBox("로그인된 사용자가 없습니다. 먼저 로컬 로그인을 수행해주세요.", "오류", "Iconx")
            return false
        }

        ; 1. 정상 조건: 쿠키 데이터가 존재하고 webPW가 있는 경우 (Headless/CDP 주입 방식)
        if (user.Has("webPW") && user["webPW"] != "" && HeadlessAutomation.CookieStorage.Count > 0) {
            if (taskType == "SessionCheck") {
                ; '선로출입관리', '승인정보불러오기' 등 XPLATFORM 구동용 세션 연동
                return this.LaunchXPlatformSession(user)
            }
            else if (taskType == "WorkLog_Create") {
                try {
                    return this.LaunchLogSession(user, "reg", "general")
                } catch as e {
                    LogDebug("[오류] 업무일지 생성 화면 이동 중 오류: " e.Message)
                    MsgBox("업무일지 생성 화면 이동 중 오류: " e.Message, "오류", "Iconx")
                    return false
                }
            }
            else if (taskType == "WorkLog_View") {
                existingLog := this._FindBrowserByElement(false)

                if existingLog {
                    ; 일지가 이미 존재 → 리스트 페이지를 먼저 준비(최대화/active)
                    if !this._ActivateListWindow() {
                        ; 리스트 창이 없으면(Case 3) 리스트만 생성
                        try this.LaunchLogSession(user, "mod", "general", true)
                    }
                    ; 기존 일지 페이지 active
                    Sleep 200
                    WinActivate("ahk_id " existingLog.BrowserId)
                    return existingLog
                }

                ; 일지가 없으면 생성 (Case 1, 2, 5)
                try {
                    return this.LaunchLogSession(user, "mod", "general")
                } catch as e {
                    LogDebug("[오류] 업무일지 조회 화면 이동 중 오류: " e.Message)
                    MsgBox("업무일지 조회 화면 이동 중 오류: " e.Message, "오류", "Iconx")
                    return false
                }
            }
            return true
        } else {
            ; 2. 폴백 (Fallback): webpw가 없는 경우 단순 레거시 런칭 방식으로 직접 제어
            if (taskType == "SessionCheck") {
                Run(
                    "https://mis.humetro.busan.kr/FS/xui/install/x_installChromeSSO.jsp?gv_selSystGubn=LA&gv_userBrowser=Edg"
                )
                if WinWait("개별업무통합관리", , 15) {
                    while !WinExist("개별업무통합관리 - 선로출입현황 조회") {
                        if (PixelGetColor(450, 470) == 0x0063B5) {
                            targetId := WinExist("A")
                            WinClose("ahk_id " targetId)

                            ; login_start
                            Run("msedge.exe https://btcep.humetro.busan.kr/user/login.face?destination=%2Fportal%2F")
                            if WinWaitActive(":: 부산교통공사", , 30) {
                                Sleep 500
                                LogDebug("[알림] 로그인이 필요합니다 (MsgBox 표시)")
                                MsgBox("로그인이 필요합니다", "알림", "icon! T5")
                            }
                            if !WinWait(":: 부산교통공사 :: ", , 15) {
                                LogDebug("[오류] 로그인 15초 타임아웃")
                                MsgBox("로그인 15초 타임아웃", "오류", "Iconx")
                                return false
                            }
                            Run(
                                "https://mis.humetro.busan.kr/FS/xui/install/x_installChromeSSO.jsp?gv_selSystGubn=LA&gv_userBrowser=Edg"
                            )
                            break
                        }
                        Sleep 500
                    }
                }

                if !WinWait("개별업무통합관리 - 선로출입현황 조회", , 30) {
                    LogDebug("[오류] 선로출입현황 조회 창 대기 타임아웃")
                    MsgBox("선로출입현황 조회 창 대기 타임아웃", "오류", "Iconx")
                    return false
                }
                return true
            }
            else if (taskType == "WorkLog_Create") {
                cUIA := this._PrepareSessionOnly(user)
                if !cUIA
                    return false

                if !this._GoToWorklogList(cUIA)
                    return false

                try {
                    cUIA.WaitElement({ LocalizedType: "링크", Name: "생성" }, 5000).Invoke()
                } catch {
                    LogDebug("[오류] 생성 버튼 대기 타임아웃")
                    MsgBox("생성 버튼 대기 타임아웃.`n매크로 동작을 중지합니다.", "오류", "Iconx")
                    return false
                }

                loop 20 {
                    if cBrowser := this._FindBrowserByElement(true)
                        return cBrowser
                    Sleep 250
                }
                LogDebug("[오류] 팝업창(업무일지 생성)을 감지하지 못함")
                MsgBox("팝업창(업무일지 생성)을 감지하지 못했습니다.", "오류", "Iconx")
                return false
            }
            else if (taskType == "WorkLog_View") {
                ; 조회 시 기존 팝업창이 열려있는지 1회 확인
                if cUIA := this._FindBrowserByElement(false)
                    return cUIA

                cUIA := this._PrepareSessionOnly(user)
                if !cUIA
                    return false
                return this._NavToWorkLogView(cUIA, user)
            }
            return true
        }
    }

    ; ==============================================================================
    ; [메서드] LaunchLogSession
    ; 설명: 쿠키 주입 후 Edge 브라우저를 실행하여 일지 작성/조회 창을 엽니다.
    ; ==============================================================================
    static LaunchLogSession(user, mode := "mod", browserMode := "general", skipPopup := false) {
        ; url 준비
        targetUrl := ""
        iljino := ""

        if (!skipPopup) {
            if (mode == "mod") {
                headless := HeadlessAutomation(true, LogDebug)
                if (!user.Has("arbpl") || user["arbpl"] == "") {
                    ;작업장코드가 없으면 inputbox로 직접 입력받음
                    arbplinput := InputBox("작업장 코드를 불러오는데 실패했습니다`n직접 입력해 주세요`n`n (호포전기분소 예 : 5129)", "작업장코드 입력",
                        "w250 h150", "5129")

                    if arbplinput.result = "Cancel"
                        return false
                    user["arbpl"] := arbplinput.Value
                }
                iljino := headless.GetTodayWorkLogNumber(user["id"], user["arbpl"])

                if (iljino == "") {
                    LogDebug("[오류] 오늘자 일지번호 찾기 실패 (작업장코드: " user["arbpl"] ")")
                    MsgBox("오늘자 일지번호 찾기에 실패했습니다`n재시작 후 다시 시도해 보기 바랍니다`n(작업장코드: " user["arbpl"] ")")
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

        ; 9222 포트가 이미 열려있는지 확인
        isConnected := false
        try {
            req := ComObject("WinHttp.WinHttpRequest.5.1")
            req.Open("GET", "http://127.0.0.1:9222/json/version", false)
            req.Send()
            if (req.Status == 200)
                isConnected := true
        } catch {
            isConnected := false
        }

        if (!isConnected) {
            try {
                Run(Format('"{1}" {2}', edgePath, args), , "Max", &edgePid)
            } catch as e {
                LogDebug("[오류] 브라우저 실행 실패: " e.Message)
                MsgBox("브라우저 실행 실패: " e.Message, "오류", "Iconx")
                return false
            }
        }

        try {
            ; Chrome 인스턴스 연결
            chromeInst := Chrome([], , , 9222)
            
            if (isConnected && chromeInst.HasProp("PID")) {
                try WinActivate("ahk_pid " chromeInst.PID)
            }

            pages := chromeInst.GetPageList()
            page := ""

            ; 팝업(일지 상세)이 아닌 페이지를 우선 선택 (과거일지 팝업 회피)
            page := this._FindNonPopupPage(pages)

            if (!page) {
                ; 비-팝업 페이지가 없으면 새 탭 생성 (기존 일지 팝업 보호)
                try {
                    newReq := ComObject("WinHttp.WinHttpRequest.5.1")
                    newReq.Open("PUT", "http://127.0.0.1:9222/json/new?about:blank", false)
                    newReq.Send()
                    if (newReq.Status == 200) {
                        newTabInfo := JSON.parse(newReq.ResponseText)
                        if newTabInfo.Has("webSocketDebuggerUrl") {
                            wsUrl := StrReplace(newTabInfo["webSocketDebuggerUrl"], "localhost", "127.0.0.1")
                            page := Chrome.Page(wsUrl)
                        }
                    }
                }
            }

            if (!page) {
                ; 최종 fallback: 아무 page 타입이라도 선택
                for p in pages {
                    if (p.Has("type") && p["type"] == "page") {
                        if (p.Has("webSocketDebuggerUrl")) {
                            wsUrl := StrReplace(p["webSocketDebuggerUrl"], "localhost", "127.0.0.1")
                            page := Chrome.Page(wsUrl)
                            break
                        }
                    }
                }
            }

            if (page) {
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

                    ; [Fetch 인터셉터 셋업]
                    sabun := user.Has("id") ? user["id"] : ""
                    arbpl := user.Has("arbpl") ? user["arbpl"] : "5129"
                    dept := user.Has("department") ? user["department"] : "호포전기분소"

                    fakeResp := ":" sabun "::5010:전기사업소:" arbpl ":" dept "::::::::::::::::::::::::::::::::"
                    fakeB64 := WebAutoLogin._Base64Encode(fakeResp)

                    fetchState := { reqId: "", done: false }

                    interceptCB(msg) {
                        if (fetchState.done)
                            return
                        if (!msg.Has("method") || msg["method"] != "Fetch.requestPaused")
                            return

                        params := msg["params"]
                        reqId := params["requestId"]
                        postData := (params.Has("request") && params["request"].Has("postData")) ? params["request"][
                            "postData"] : ""

                        if InStr(postData, "BOOKSCH") {
                            fetchState.reqId := reqId
                        } else {
                            try page.Call("Fetch.continueRequest", Map("requestId", reqId), false)
                        }
                    }

                    page._callback := interceptCB
                    page.Call("Fetch.enable", Map("patterns", [Map("urlPattern", "*WorkLogData*", "requestStage",
                        "Request")]))

                    ; 업무일지 리스트 페이지 이동 (AJAX 발생)
                    page.Call("Page.navigate", Map("url", this.Worklog_List)) ;헤더 포함

                    loop 150 {
                        Sleep 100
                        if (fetchState.reqId == "")
                            continue

                        try {
                            page.Call("Fetch.fulfillRequest", Map(
                                "requestId", fetchState.reqId,
                                "responseCode", 200,
                                "responseHeaders", [Map("name", "content-type", "value", "text/plain; charset=utf-8")],
                                "body", fakeB64
                            ))
                        }

                        fetchState.done := true
                        fetchState.reqId := ""
                        break
                    }

                    page.Call("Fetch.disable")
                    page._callback := 0

                    ; 리스트 페이지 로딩 대기 및 최대화
                    WinWait("업무일지관리 - 부산교통공사", , 5)
                    try {
                        listHwnd := WinExist("업무일지관리 - 부산교통공사")
                        if listHwnd {
                            WinMaximize("ahk_id " listHwnd)
                            WinActivate("ahk_id " listHwnd)
                        }
                    }

                    ; skipPopup이면 리스트만 준비하고 종료
                    if (skipPopup)
                        return true

                    modeCode := (mode == "mod") ? "MOD" : "REG"

                    js := "let cnt = 0; "
                        . "let inter = setInterval(() => { "
                        . "    try {"
                        . "        const win = document"
                        . "            .querySelector('#ivuFrm_page0ivu1')"
                        . "            .contentDocument"
                        . "            .querySelector(`"iframe[name='isolatedWorkArea']`")"
                        . "            .contentWindow; "
                        . "        if (win.$ && win.fn_detail_open && win.document?.WorkLogForm) {"
                        . "            clearInterval(inter); "

                    if (mode == "mod" && iljino != "")
                        js .= "            win.$('#V_ILJINO').val('" iljino "'); "

                    js .= "            win.$('#I_MODE').val('" modeCode "'); "
                        .
                        "            win.fn_detail_open('WorkLogForm', 'reg', 'kr.busan.humetro.cbo.erp.work_log.WorkLogReg', 'width=1024,height=760,scrollbars=yes,resizable=yes,status=yes'); "
                        . "        } "
                        . "    } catch (e) {}"
                        . "    if (cnt++ > 20) clearInterval(inter); "
                        . "}, 100); "

                    page.Evaluate(js)
                }
            }
        } catch as e {
            LogDebug("[오류] 브라우저 제어(쿠키/이동) 오류: Line" e.Line " " e.Message)
            MsgBox("브라우저 제어(쿠키/이동) 오류: Line" e.Line "`n" e.Message, "오류", "Iconx")
            return false
        }

        ; 열린 팝업창(일지 상세창)의 UIA 객체를 획득하여 반환

        ; JS(fn_detail_open) 실행 후 창이 2개(메인, 팝업)가 될 때까지 최대 15초 대기
        loop 30 {
            Sleep 100
            if cUIA := this._FindBrowserByElement()
                return cUIA
        }
        LogDebug("[오류] 일지 상세창(팝업) 호출 시간 초과 - cUIA 연결 실패")
        MsgBox("일지 상세창(팝업) 호출 시간 초과`ncUIA 연결 실패", "알림", "Iconx")
        return false
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
                LogDebug("[오류] XPLATFORM ep 쿠키 정보 없음")
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
            LogDebug("[오류] XPLATFORM 실행 중 오류: " err.Message)
            MsgBox("XPLATFORM 실행 중 오류 발생 (런처 데몬이 종료되었거나 포트가 다릅니다): " err.Message, "오류", "Iconx")
            return false
        }
    }

    ; ==============================================================================
    ; [내부] _FindBrowserByElement (단순 레거시 복원)
    ; 설명: 현재 띄워진 브라우저 창들 중 "부산교통공사" 타이틀이면서 당일자 일지가 열린 탭을 탐색
    ; ==============================================================================
    static _FindBrowserByElement(create := false) {
        targetBrowsers := ["msedge.exe", "chrome.exe", "whale.exe"]
        for exe in targetBrowsers {
            if !ProcessExist(exe)
                continue
            try hwndList := WinGetList("ahk_exe " exe)
            catch
                continue

            for hwnd in hwndList {
                if !InStr(WinGetTitle("ahk_id " hwnd), "부산교통공사 - ")
                    continue
                try {
                    cUIA := UIA_Browser("ahk_id " hwnd)
                    nowDate := FormatTime(DateAdd(A_Now, create ? 0 : -9, "Hours"), "yyyy-MM-dd")
                    if cUIA.FindElement({ AutomationId: "I_GIJUND", Value: nowDate }) {
                        WinRestore("ahk_id " hwnd)
                        WinActivate("ahk_id " hwnd)
                        return cUIA
                    }
                }
            }
        }
        return false
    }

    ; ==============================================================================
    ; [내부] _ActivateListWindow
    ; 설명: "업무일지관리" 타이틀 윈도우를 찾아 최대화 및 활성화
    ; ==============================================================================
    static _ActivateListWindow() {
        targetBrowsers := ["msedge.exe", "chrome.exe", "whale.exe"]
        for exe in targetBrowsers {
            if !ProcessExist(exe)
                continue
            try hwndList := WinGetList("ahk_exe " exe)
            catch
                continue
            for hwnd in hwndList {
                if InStr(WinGetTitle("ahk_id " hwnd), "업무일지관리 - 부산교통공사") {
                    WinMaximize("ahk_id " hwnd)
                    WinActivate("ahk_id " hwnd)
                    return true
                }
            }
        }
        return false
    }

    ; ==============================================================================
    ; [내부] _FindNonPopupPage
    ; 설명: CDP 페이지 목록에서 팝업(일지 상세)이 아닌 페이지를 우선 선택
    ;       리스트 페이지 > 일반 탭 순으로 우선순위
    ; ==============================================================================
    static _FindNonPopupPage(pages) {
        ; 1차: 리스트 페이지(workset.03) 우선 탐색
        for p in pages {
            if (p.Has("type") && p["type"] == "page" && p.Has("url")) {
                if (InStr(p["url"], "workset.03") || InStr(p["url"], "WorkLogList")) {
                    if (p.Has("webSocketDebuggerUrl")) {
                        wsUrl := StrReplace(p["webSocketDebuggerUrl"], "localhost", "127.0.0.1")
                        return Chrome.Page(wsUrl)
                    }
                }
            }
        }
        ; 2차: 팝업(WorkLogReg = 일지 상세)이 아닌 일반 탭 선택
        for p in pages {
            if (p.Has("type") && p["type"] == "page" && p.Has("url")) {
                if InStr(p["url"], "WorkLogReg")
                    continue
                if (p.Has("webSocketDebuggerUrl")) {
                    wsUrl := StrReplace(p["webSocketDebuggerUrl"], "localhost", "127.0.0.1")
                    return Chrome.Page(wsUrl)
                }
            }
        }
        return false
    }

    ; ==============================================================================
    ; [내부] _PrepareSessionOnly
    ; ==============================================================================
    static _PrepareSessionOnly(user) {
        exeName := this.GetBrowserExe()
        Run(exeName " " this.PortalURL)
        WinMaximize hwnd := WinWaitActive("ahk_exe " exeName, , 5)

        try {
            cUIA := UIA_Browser(hwnd)
        } catch as e {
            LogDebug("[오류] cUIA 연결 실패 (_PrepareSessionOnly)")
            MsgBox("cUIA 연결 실패", "오류", "Iconx")
            return false
        }

        ; IsLoggedIn 5초 대기 (50회 x 100ms)
        if !this.IsLoggedIn(cUIA, true, 50) {
            return this.Login(user["id"], user.Has("webPW") ? user["webPW"] : "", user.Has("pw2") ? user["pw2"] : "",
            cUIA)
        }
        return cUIA
    }

    ; ==============================================================================
    ; [내부] _NavToWorkLogView
    ; ==============================================================================
    static _NavToWorkLogView(cUIA, user) {
        try {
            if !this._GoToWorklogList(cUIA)
                return false
            targetDate := FormatTime(DateAdd(A_Now, -9, "Hours"), "yyyyMMdd")
            dept := user.Has("department") ? user["department"] : "호포전기분소"
            targetName := targetDate " " dept " 업무일지"

            try {
                ; 페이지 진입 직후 5초 대기
                cUIA.WaitElement({ LocalizedType: "텍스트", Name: targetName }, 5000).Click("Left")
                Sleep 250
                ; 기본 2초 대기
                cUIA.WaitElement({ LocalizedType: "링크", Name: "변경/조회" }, 2000).Invoke()
            } catch {
                LogDebug("[오류] 오늘자 업무일지 요소 대기 타임아웃: " targetName)
                MsgBox("오늘자 업무일지(" targetName ") 요소 대기 타임아웃.`n매크로 동작을 중단합니다.", "오류", "Icon!")
                return false
            }

            loop 20 {
                if cBrowser := this._FindBrowserByElement(false)
                    return cBrowser
                Sleep 250
            }
            LogDebug("[오류] 업무일지 조회 팝업창을 찾을 수 없음")
            MsgBox("업무일지 조회 팝업창을 찾을 수 없습니다.", "오류", "Iconx")
            return false
        } catch as e {
            LogDebug("[오류] 업무일지 조회 화면 이동 중 오류: " e.Message)
            MsgBox("업무일지 조회 화면 이동 중 오류: " e.Message, "오류", "Iconx")
            return false
        }
    }

    ; ==============================================================================
    ; [헬퍼] 업무일지 리스트 메뉴 이동
    ; ==============================================================================
    static _GoToWorklogList(cUIA) {
        currentTitle := WinGetTitle("ahk_id " cUIA.BrowserId)
        isList := InStr(currentTitle, "업무일지관리 - 부산교통공사")
        isERP := InStr(currentTitle, "ERP포털시스템 - 부산교통공사")
        isIntegrated := InStr(currentTitle, ":: 부산교통공사 ::")

        if (isList)
            return true

        if (!isERP) {
            if (isIntegrated) {
                try {
                    cUIA.WaitElement({ Type: "Link", Name: "ERP" }, 2000).Invoke()
                } catch {
                    cUIA.navigate(this.ERP_PortalURL, , 10000)
                }
            } else {
                cUIA.navigate(this.ERP_PortalURL, , 10000)
            }

            try {
                foundEl := cUIA.WaitElement([{ AutomationId: "userId" }, { Name: "업무일지 업무일지" }], 10000)
                if (foundEl.AutomationId == "userId") {
                    user := ConfigManager.CurrentUser
                    if this.Login(user["id"], user.Has("webPW") ? user["webPW"] : "", user.Has("pw2") ? user["pw2"] :
                        "", cUIA) {
                        cUIA.navigate(this.ERP_PortalURL, , 10000)
                        cUIA.WaitElement({ Name: "업무일지 업무일지" }, 10000)
                    } else {
                        throw Error("세션 만료 후 재로그인 실패")
                    }
                }
            } catch as e {
                LogDebug("[오류] 페이지 이동 중 타임아웃: " e.Message)
                MsgBox("페이지 이동 중 타임아웃 오류 발생: " e.Message, "오류", "Iconx")
                return false
            }
        }

        try {
            cUIA.WaitElement({ Name: "업무일지 업무일지" }, 10000).Invoke()
        } catch {
            LogDebug("[오류] 업무일지 메뉴 타임아웃")
            MsgBox("업무일지 메뉴 타임아웃.`n매크로 동작을 중단합니다.", "오류", "Iconx")
            return false
        }

        return true
    }

    ; ============================================================
    ; 정적 헬퍼 메서드 추가
    ; ============================================================

    static _Base64Encode(str) {
        bytes := Buffer(StrPut(str, "UTF-8") - 1)
        StrPut(str, bytes, "UTF-8")
        size := 0
        DllCall("Crypt32\CryptBinaryToString", "Ptr", bytes, "UInt", bytes.Size, "UInt", 0x40000001, "Ptr", 0, "UInt*", &
            size)
        buf := Buffer(size * 2)
        DllCall("Crypt32\CryptBinaryToString", "Ptr", bytes, "UInt", bytes.Size, "UInt", 0x40000001, "Ptr", buf,
            "UInt*", &size)
        return StrReplace(StrReplace(StrGet(buf, "UTF-16"), "`r", ""), "`n", "")
    }

    static _BuildAlertPatch() {
        js := "
        (
        (function() {
            ; ── 진단용 마커: 이 스크립트가 실행된 window에 표시
            window.__alertPatched = true;
            console.log('[PATCH] addScriptToEvaluateOnNewDocument 적용됨:', location.href);

            ; ── alert 오버라이드
            var _orig = window.alert;
            window.alert = function(msg) {
                if (typeof msg === 'string' && (
                    msg.includes('플랜트')   ||
                    msg.includes('작업장')   ||
                    msg.includes('일치')     ||
                    msg.includes('필수항목') ||
                    msg.includes('부서') )) {
                    console.warn('[PATCH] alert 억제:', location.href, '|', msg);
                    return;
                }
                return _orig.apply(this, arguments);
            };
        })();
        )"

        return js
    }

    static _BuildPatchScript(user) {
        patchJS := "
        (
        (function() {

            // 원본 alert 함수 저장
            var _origAlert = window.alert;

            // alert 함수 가로채기
            window.alert = function(msg) {
                // 특정 메시지 포함된 alert은 차단
                if (typeof msg === 'string' && (
                    msg.includes('플랜트') ||
                    msg.includes('작업장') ||
                    msg.includes('일치') ||
                    msg.includes('부서'))) {
                    console.warn('[패치] alert 차단:', msg);
                    return;
                }
                // 다른 alert은 원본 alert으로 처리
                return _origAlert.apply(this, arguments);
            };

            // #ivuFrm_page0ivu1 안의 #isolatedWorkArea 내에서도 alert 차단
            var iframe = document.querySelector('#ivuFrm_page0ivu1');
            if (iframe) {
                var iframeWindow = iframe.contentWindow;
                // iframe 내에서도 alert 가로채기
                var _iframeAlert = iframeWindow.alert;
                iframeWindow.alert = function(msg) {
                    if (typeof msg === 'string' && (
                        msg.includes('플랜트') ||
                        msg.includes('작업장') ||
                        msg.includes('일치') ||
                        msg.includes('부서'))) {
                        console.warn('[패치] iframe 내 alert 차단:', msg);
                        return;
                    }
                    return _iframeAlert.apply(this, arguments);
                };
            }

            // #isolatedWorkArea 내의 alert 차단
            var isolatedWorkArea = document.querySelector('#isolatedWorkArea');
            if (isolatedWorkArea) {
                var _isolatedAlert = isolatedWorkArea.contentWindow.alert;
                isolatedWorkArea.contentWindow.alert = function(msg) {
                    if (typeof msg === 'string' && (
                        msg.includes('플랜트') ||
                        msg.includes('작업장') ||
                        msg.includes('일치') ||
                        msg.includes('부서'))) {
                        console.warn('[패치] isolatedWorkArea 내 alert 차단:', msg);
                        return;
                    }
                    return _isolatedAlert.apply(this, arguments);
                };
            }

        })();
        )"

        return patchJS
    }

    static _BuildPatchScript__(user) {
        ; AHK의 user 객체 값을 JS에 직접 삽입
        ; → SESS_* 서버 렌더링에 의존하지 않아도 됨
        arbpl := user["arbpl"]   ; 예: "5129"
        pernr := user["id"]      ; 예: "116713"

        patchJS := "
        (
        (function() {

            /* ── ① alert 억제 ─────────────────────────────────────
               fn_Sch2() 내부의 '플랜트가 작업장 플랜트와 일치하지 않습니다'
               alert이 실행 흐름을 막는 것을 방지                          */
            var _origAlert = window.alert;
            window.alert = function(msg) {
                if (typeof msg === 'string' && (
                    msg.includes('플랜트') ||
                    msg.includes('작업장') ||
                    msg.includes('일치') ||
                    msg.includes('부서') )) {
                    console.warn('[패치] alert 억제:', msg);
                    return;   // 무시
                }
                return _origAlert.apply(this, arguments);
            };

            /* ── ② BOOKSCH XHR 인터셉트 ───────────────────────────
               즐겨찾기 없는 응답이 돌아오면:
                 a) jQuery success 콜백(else 분기 + fn_Sch2)이 먼저 실행됨
                 b) alert은 ①에서 억제됨
                 c) loadend 후 200ms 뒤 값을 교정하고 fn_Sch2() 재실행   */
            var _arbpl  = '__ARBPL__';    ; AHK가 치환
            var _pernr  = '__PERNR__';    ; AHK가 치환

            var _origSend = XMLHttpRequest.prototype.send;
            XMLHttpRequest.prototype.send = function(body) {
                if (body && typeof body === 'string' && body.includes('BOOKSCH')) {
                    var xhr = this;
                    xhr.addEventListener('loadend', function() {
                        try {
                            var resp = xhr.responseText || '';

                            /* 첫 글자가 ':'가 아니면 즐겨찾기 없는 실패 응답 */
                            if (resp.charAt(0) !== ':') {

                                setTimeout(function() {

                                    /* ARBPL이 비어있거나 디폴트(1081)로 세팅된 경우만 교정 */
                                    var curArbpl = (typeof $ !== 'undefined')
                                                   ? $('#I_ARBPL').val() : '';
                                    if (curArbpl !== '' && curArbpl !== '1081') return;

                                    /* SESS_* (서버 렌더링) 또는 AHK 주입값 우선 사용 */
                                    var targetWerks    = (typeof SESS_WERKS    !== 'undefined' && SESS_WERKS)
                                                         ? SESS_WERKS    : '5010';
                                    var targetArbpl    = (typeof SESS_ARBPL    !== 'undefined' && SESS_ARBPL)
                                                         ? SESS_ARBPL    : _arbpl;
                                    var targetArbplTxt = (typeof SESS_ARBPLTEXT !== 'undefined' && SESS_ARBPLTEXT)
                                                         ? SESS_ARBPLTEXT : '';

                                    if (!targetArbpl) return;  ; 값이 없으면 포기

                                    /* 필드 교정 */
                                    $('#I_ARWRK').val(targetWerks);
                                    $('#I_WERKS').val(targetWerks);
                                    $('#I_ARBPL').val(targetArbpl);
                                    $('#I_ARBPLTEXT').val(targetArbplTxt);

                                    console.log('[패치] 작업장 교정 완료:', targetArbpl, targetArbplTxt);

                                    /* fn_Sch2() 재실행 → 이번엔 ARWRK==WERKS이므로 정상 검색 */
                                    if (typeof fn_Sch2 === 'function') fn_Sch2();

                                }, 200);
                            }
                        } catch(e) {
                            console.warn('[패치] BOOKSCH 처리 오류:', e);
                        }
                    });
                }
                return _origSend.apply(this, arguments);
            };

        })();
        )"
        ; AHK 변수 치환
        patchJS := StrReplace(patchJS, "__ARBPL__", arbpl)
        patchJS := StrReplace(patchJS, "__PERNR__", pernr)
        return patchJS
    }

    ; ==============================================================================
    ; [헬퍼] 브라우저 이름 탐색
    ; ==============================================================================
    static GetBrowserExe() {
        targetBrowsers := ["msedge.exe", "chrome.exe", "whale.exe"]
        for exe in targetBrowsers {
            if !ProcessExist(exe)
                continue
            try {
                hwndList := WinGetList("ahk_exe " exe)
            } catch {
                continue
            }
            for hwnd in hwndList {
                if !InStr(WinGetTitle(hwnd), "부산교통공사")
                    continue
                try {
                    cUIA := UIA_Browser("ahk_id " hwnd)
                    tabs := cUIA.GetAllTabs()
                    for tabItem in tabs {
                        if InStr(tabItem.Name, "부산교통공사")
                            return exe
                    }
                } catch {
                    continue
                }
            }
        }
        return "msedge.exe"
    }

    ; ==============================================================================
    ; [메서드] IsLoggedIn
    ; ==============================================================================
    static IsLoggedIn(cUIA := "", silent := false, loops := 30) {
        try {
            if !cUIA
                cUIA := UIA_Browser("A")
            loop loops {
                try {
                    if cUIA.FindElement({ Name: "로그아웃" })
                        return cUIA
                }
                try {
                    if cUIA.FindElement({ Name: "업무일지 업무일지" })
                        return cUIA
                }
                try {
                    if cUIA.FindElement({ AutomationId: "userId" })
                        return false
                }
                Sleep 100
            }
            if !silent
                LogDebug("[오류] 로그인 상태 확인 실패 또는 알 수 없는 상태")
                MsgBox("로그인 상태 확인을 지연(로딩 중)하거나 알 수 없는 상태입니다.", "알림", "Iconi")
            return false
        } catch {
            return false
        }
    }

    ; ==============================================================================
    ; [메서드] Login
    ; ==============================================================================
    static Login(id, pw, pw2, cUIA := "") {
        if !cUIA
            cUIA := UIA_Browser("A")

        if (id == "" || pw == "") {
            LogDebug("[알림] 통합pw 미설정 - 수동 로그인 대기")
            MsgBox("통합pw가 지정되어 있지 않습니다. 브라우저에서 직접 로그인해주세요.`n확인을 누르면 15초간 로그인을 대기합니다.", "로그인 대기", "Iconi")
            loop 15 {
                Sleep 1000
                if (this.IsLoggedIn(cUIA, true, 10))
                    return cUIA
            }
            LogDebug("[오류] 로그인 대기 시간(15초) 초과")
            MsgBox("로그인 대기 시간(15초)이 초과되었습니다.`n위 기능 사용이 제한되며 매크로 작업을 종료합니다.", "시간 초과", "Iconx")
            return false
        }

        try {
            cUIA.WaitElement({ AutomationId: "userId" }, 2000).Value := id
            cUIA.FindElement({ AutomationId: "password" }).Value := pw
            cUIA.FindElement({ ClassName: "btn_login" }).Invoke()

            try {
                cUIA.WaitElement({ AutomationId: "certi_num" }, 2000).Value := pw2
                cUIA.FindElement({ ClassName: "btn_blue" }).Invoke()
            } catch as e {
                LogDebug("[오류] 2차인증 실패: " e.Message)
                MsgBox("2차인증 실패`n" e.Message, "오류")
                return false
            }

            try {
                cUIA.WaitElement({ Name: "로그아웃" }, 10000)
                Sleep 1000
                return cUIA
            } catch {
                LogDebug("[오류] 로그인 후 응답 지연 또는 실패")
                MsgBox("로그인 후 응답이 지연되거나 실패했습니다.", "오류")
                return false
            }
        } catch as e {
            LogDebug("[오류] 자동 로그인 실패: " e.Message)
            MsgBox("자동 로그인 실패: " e.Message, "오류", "Iconx")
            return false
        }
    }
}
