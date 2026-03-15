#Requires AutoHotkey v2.0
#SingleInstance Force
#Include "설정관리.ahk"
#Include "Lib\JSON.ahk"
#Include "Lib\WebView2.ahk"
#Include "Lib\CryptAES.ahk"
#Include "Lib\UIA.ahk"
#Include "Lib\UIA_Browser.ahk"
#Include "Lib\Chrome.ahk"
#Include "단축기능.ahk"
#Include "유저로그인.ahk"
#Include "URL.ahk"
#Include "선로출입.ahk"
#Include "차량일지.ahk"
#Include "업무일지.ahk"
#Include "ERP점검.ahk"
#Include "헤드리스.ahk"

; ==============================================================================
; 컴파일러 지시문
; ==============================================================================
;@Ahk2Exe-SetVersion 3.2.4.0
;@Ahk2Exe-SetProductVersion v3.2.4
;@Ahk2Exe-SetDescription 통합자동화
; ==============================================================================
; ==============================================================================
; 초기화
; ==============================================================================
global AppVersion := "v3.2.4"
global wvc := ""
global wv := ""
global MainGui := ""
global LoadingGui := ""

if !ConfigManager.Load()
    ExitApp

; 리로드 시 기존 로그인 쿠키(세션) 환경 복원
HeadlessAutomation.LoadCookieStorage()

; 복호화 실패로 암호가 초기화되었을 경우 알림 메시지 출력
if (ConfigManager.NeedsPasswordReset) {
    MsgBox("프로그램 위치 변경 또는 복사가 감지되어`n저장된 비밀번호가 초기화되었습니다.`n`n[설정] > [내 정보] 메뉴에서 비밀번호를 다시 입력해주세요.", "보안 알림", "Iconi")
}

; ==============================================================================
; [Self-Bootstrapping] 필수 파일 내장 및 추출
; ==============================================================================
; 설명: 컴파일된 Main.exe는 아래 파일들을 내장하고 있으며, 실행 시마다 최신 버전으로 덮어씁니다.
; 이를 통해 Main.exe만 배포해도 모든 구성요소가 자동으로 업데이트됩니다.
if (A_IsCompiled) {
    try {
        ; 1. 실행 파일 (updater.exe)
        FileInstall("Updater.exe", A_ScriptDir "\Updater.exe", 1)

        ; 2. 라이브러리 (DLL)
        if !DirExist(A_ScriptDir "\Lib\64bit")
            DirCreate(A_ScriptDir "\Lib\64bit")
        FileInstall("Lib\64bit\WebView2Loader.dll", A_ScriptDir "\Lib\64bit\WebView2Loader.dll", 1)

        ; 3. UI 리소스
        if !DirExist(A_ScriptDir "\ui\img")
            DirCreate(A_ScriptDir "\ui\img")
        if !DirExist(A_ScriptDir "\ui\assets")
            DirCreate(A_ScriptDir "\ui\assets")

        FileInstall("ui\index.html", A_ScriptDir "\ui\index.html", 1)
        FileInstall("ui\app.js", A_ScriptDir "\ui\app.js", 1)
        FileInstall("ui\style.css", A_ScriptDir "\ui\style.css", 1)
        FileInstall("ui\img\loading.gif", A_ScriptDir "\ui\img\loading.gif", 1)
        FileInstall("ui\assets\icon.ico", A_ScriptDir "\ui\assets\icon.ico", 1)

        ;4. UI 라이브러리
        if !DirExist(A_ScriptDir "\ui\lib")
            DirCreate(A_ScriptDir "\ui\lib")
        FileInstall("ui\lib\vue.global.prod.js", A_ScriptDir "\ui\lib\vue.global.prod.js", 1)

    } catch as e {
        ; 파일 사용 중 등으로 실패할 경우 로그만 남기고 진행 (치명적이지 않음)
        LogDebug("내장 파일 추출 경고: " e.Message)
    }
}

; 업데이트 확인 (차단 대기) - 컴파일된 경우(프로덕션 모드)에만 실행
skipUpdate := false
for arg in A_Args {
    if (arg = "/skipupdate")
        skipUpdate := true
}

if (A_IsCompiled && !skipUpdate) {
    updaterExe := A_ScriptDir "\Updater.exe"
    if FileExist(updaterExe) {
        try RunWait(updaterExe)
    }
}

; ------------------------------------------------------------------------------
; 초기화 시퀀스 (비동기 호출)
; ------------------------------------------------------------------------------
PerformReadySequence() {

    profiles := ConfigManager.GetProfiles()
    payload := Map("type", "initLogin", "users", profiles)
    wv.PostWebMessageAsJson(JSON.stringify(payload))

    ; 자동 로그인 복구 초기화
    restartFile := A_ScriptDir "\.restart_login"
    if FileExist(restartFile) {
        try {
            savedID := FileRead(restartFile)
            FileDelete(restartFile)
            if (savedID != "") {
                ; 복구 로그인 시에도 AHK 내부 상태 동기화 (아래의 로그인 처리와 동일)
                ConfigManager.Set("appSettings.lastUser", savedID)

                userRoot := ConfigManager.GetUserRoot(savedID)
                profile := userRoot.Has("profile") ? userRoot["profile"] : Map("id", savedID, "name", "Unknown")
                ConfigManager.CurrentUser := profile
                LoadConfigData()

                payload := Map("type", "loginSuccess", "profile", profile)

                jsonResp := JSON.stringify(payload)
                wv.PostWebMessageAsJson(jsonResp)

                ; [추가] 근무 조 정보 전송 (복구 시에도 업데이트)
                userTeam := profile.Has("team") ? profile["team"] : ""
                shiftContext := WorkLogManager.GetCurrentContext(userTeam)
                shiftPayload := Map("type", "updateShiftStatus", "data", shiftContext)
                wv.PostWebMessageAsJson(JSON.stringify(shiftPayload))

                ; UI 상태 복원 (로그인 복구 후 실행)
                stateFile := A_ScriptDir "\.restore_state.json"
                if FileExist(stateFile) {
                    try {
                        jsonStr := FileRead(stateFile)
                        FileDelete(stateFile)
                        if (jsonStr != "") {
                            state := JSON.parse(jsonStr)
                            payload := Map("type", "restoreUiState", "data", state)
                            wv.PostWebMessageAsJson(JSON.stringify(payload))
                        }
                    }
                }

                ; 기존 쿠키가 살아있다면 버튼 활성화 상태 복구 (headlessReady 알림)
                if (HeadlessAutomation.CookieStorage.Count > 0) {
                    wv.PostWebMessageAsJson(JSON.stringify(Map("type", "headlessReady")))
                }

                ; [추가] 작업보고.sap 파일 갱신 (자동 로그인)
                UpdateSapFile(savedID)

                ; [추가] 자동 종료 타이머 시작
                StartAutoExitTimer(profile)
            }
        }
    }

    ; ERP 상태 폴링 시작
    ERP점검.StartPolling()
}

; ==============================================================================
; 유틸리티 함수
; ==============================================================================
LogDebug(text) {
    try {
        FileAppend(FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") " [Main] " text "`n", A_ScriptDir "\debug_main.txt",
        "UTF-8")
    }
}

UnzipFile(zipPath, destDir) {
    shell := ComObject("Shell.Application")
    if !DirExist(destDir)
        DirCreate(destDir)

    try {
        zipFolder := shell.NameSpace(zipPath)
        destFolder := shell.NameSpace(destDir)

        if (zipFolder && destFolder) {
            ; 4 (No Progress UI) + 16 (Yes to All)
            destFolder.CopyHere(zipFolder.Items, 20)
            Sleep 1000
        }
    } catch as e {
        throw Error("Unzip failed: " e.Message)
    }
}

LogDebug("========== Main.ahk 시작 ==========")

; ==============================================================================
; 창 및 트레이 도우미
; ==============================================================================
OnWindowClose(*) {
    if (ConfigManager.GetCurrentUserID() == "") {
        OnExitApp()
    }
    MainGui.Hide()
    TrayTip "시스템 트레이에서 실행 중입니다", "통합자동화", 1
}

RestoreWindow(*) {
    MainGui.Show()
    WinActivate("ahk_id " MainGui.Hwnd)
}

; 2. 메인 창
ShowMainWindow()

; ==============================================================================
; 메인 로직
; ==============================================================================
ShowMainWindow() {
    global MainGui, wvc, wv

    ; 테두리 없는 창 생성
    ; 크기 조절 가능한 테두리 없는 창 생성 (WS_THICKFRAME = +Resize, 0x00040000)
    ; 이를 통해 캡션 없이도 표준 Windows 10/11 그림자 효과를 사용할 수 있습니다.
    ; 필요한 경우 OnMessage를 통해 수동으로 크기 조절을 방지할 수 있지만, 현재는 최신 앱 동작에 따라 허용합니다.
    ; 인수 확인
    titleSuffix := ""
    for arg in A_Args {
        if (arg = "/offline")
            titleSuffix := " - 오프라인 모드"
    }

    MainGui := Gui("-Caption +Resize", "통합자동화 " AppVersion titleSuffix)
    MainGui.SetFont("S10", "Malgun Gothic")
    MainGui.BackColor := "FFFFFF"

    ; 이벤트 핸들러
    MainGui.OnEvent("Close", OnWindowClose)
    MainGui.OnEvent("Size", OnGuiSize)

    ; 트레이 메뉴 설정
    A_TrayMenu.Delete() ; 기본값 지우기
    A_TrayMenu.Add("열기", RestoreWindow)
    A_TrayMenu.Add("종료", OnExitApp)
    A_TrayMenu.Default := "열기"
    A_TrayMenu.ClickCount := 1 ; 한 번 클릭으로 복원

    ; 초기 단축키 설정 및 데이터 로드
    LoadConfigData()

    ; WebView2 설정
    try {
        dllPath := A_ScriptDir "\Lib\" (A_PtrSize = 8 ? "64bit" : "32bit") "\WebView2Loader.dll"
        if !FileExist(dllPath)
            throw Error("WebView2Loader.dll not found")

        wvc := WebView2.create(MainGui.Hwnd, , , , , , dllPath)
        wvc.IsVisible := true
        wv := wvc.CoreWebView2

        ; 브라우저 기능 비활성화
        settings := wv.Settings
        settings.AreDefaultContextMenusEnabled := false
        settings.IsZoomControlEnabled := false
        settings.IsStatusBarEnabled := false
        settings.AreBrowserAcceleratorKeysEnabled := true ; 디버그를 위해 F12 허용

        wv.add_WebMessageReceived(OnWebMessage)

        ; HTML 경로 로직
        htmlPath := A_ScriptDir "\ui\index.html"
        if !FileExist(htmlPath) {
            throw Error("HTML file not found: " htmlPath)
        }

        ; 경로 정규화
        htmlPath := StrReplace(htmlPath, "\", "/")

        ; UNC 대 로컬 경로 처리
        if (SubStr(htmlPath, 1, 2) == "//") {
            ; UNC 경로: //server/share -> file://server/share
            uri := "file:" htmlPath
        } else {
            ; 로컬 경로: C:/... -> file:///C:/...
            uri := "file:///" htmlPath
        }

        wv.Navigate(uri)

    } catch as e {
        MsgBox("Error: " e.Message)
        ExitApp
    }

    MainGui.Show("w800 h600 Center")

    if (wvc) {
        wvc.Fill()
    }

    ; DWM 확장 프레임을 사용하여 그림자 처리 (CS_DROPSHADOW나 단순 NCCALCSIZE보다 잘 작동함)
    if (VerCompare(A_OSVersion, "6.0") >= 0) {
        MARGINS := Buffer(16, 0)
        NumPut("Int", 1, MARGINS, 0) ; 왼쪽
        NumPut("Int", 1, MARGINS, 4) ; 오른쪽
        NumPut("Int", 1, MARGINS, 8) ; 위쪽
        NumPut("Int", 1, MARGINS, 12) ; 아래쪽
        DllCall("Dwmapi\DwmExtendFrameIntoClientArea", "Ptr", MainGui.Hwnd, "Ptr", MARGINS)
    }
}

OnGuiSize(guiObj, minMax, width, height) {
    if (wvc) {
        wvc.Fill()
    }
}

OnExitApp(*) {
    ; 1. UI 복구 파일 삭제
    stateFile := A_ScriptDir "\.restore_state.json"
    loginFile := A_ScriptDir "\.restart_login"
    erpTempFile := A_ScriptDir "\erp_status_temp.json"
    tempfile := A_WorkingDir . "\temp_*.json"
    try FileDelete(tempfile)
    try FileDelete(stateFile)
    try FileDelete(loginFile)
    try FileDelete(erpTempFile)

    ; 백그라운드 매니저 정리
    BackgroundProcessManager.Cleanup()

    ; 4. 종료
    ExitApp
}

; ==============================================================================
; 브리지: JS -> AHK
; ==============================================================================
; ------------------------------------------------------------------------------
; WebView 메시지 수신 핸들러 (JS -> AHK)
; 설명: WebView(React/JS)에서 window.chrome.webview.postMessage로 보낸 JSON을 처리합니다.
; ------------------------------------------------------------------------------
OnWebMessage(sender, args) {
    jsonStr := args.WebMessageAsJson

    if (jsonStr == "")
        return

    ; JSON 데이터 파싱
    try {
        msg := JSON.parse(jsonStr)
    } catch as e {
        LogDebug("WebMessage JSON 파싱 실패: " e.Message)
        return
    }

    ; 커맨드 추출
    command := msg.Has("command") ? msg["command"] : "Unknown"
    LogDebug("WebMessage 수신: " command)

    ; --- 1. 초기화 및 로그인 ---
    if (command == "ready") {
        ; [UX 개선] 초기화 로직을 비동기(Timer)로 분리하여 WebView 응답성 확보
        ; 바로 실행 시 Unzip 등으로 인해 UI 스레드가 차단될 수 있음
        SetTimer PerformReadySequence, -10
    }
    else if (command == "tryLogin") { ; 로그인 처리
        uid := msg.Has("id") ? msg["id"] : ""

        ConfigManager.Set("appSettings.lastUser", uid)

        LogDebug("로그인 프로세스 시작. ID: " uid)

        userRoot := ConfigManager.GetUserRoot(uid)
        profile := userRoot.Has("profile") ? userRoot["profile"] : Map("id", uid, "name", "Unknown")
        ConfigManager.CurrentUser := profile
        LoadConfigData()

        ; [Headless] 백그라운드 자동화 실행 (프로필 정보 활용)
        try {
            runWebPW := msg.Has("pw") ? msg["pw"] : (profile.Has("webPW") ? profile["webPW"] : "")
            runPW2 := msg.Has("pw2") ? msg["pw2"] : (profile.Has("pw2") ? profile["pw2"] : "")

            if (runWebPW != "" && runPW2 != "") {
                LogDebug("세션BG: 로그인 프로세스 시작... (ID: " uid ")")

                ; 세션BG 실행 경로 설정
                exePath := A_ScriptDir "\세션BG.exe"
                ; FileInstall은 Main이 컴파일될 때 포함되도록 함
                if A_IsCompiled and !FileExist(exePath)
                    FileInstall("세션BG.exe", exePath, 1)

                ; 기존 프로세스 정리
                BackgroundProcessManager.Cleanup()

                ; 세션BG 실행 (결과는 OnSessionBGOutput에서 처리하도록 콜백 전달)
                cmdLine := Format('"{1}" "{2}" "{3}" "{4}"', exePath, uid, runWebPW, runPW2)
                if (BackgroundProcessManager.Launch(cmdLine, OnSessionBGOutput)) {
                    LogDebug("세션BG 프로세스 실행 성공")
                } else {
                    LogDebug("세션BG 프로세스 실행 실패")
                    MsgBox("로그인 프로세스 실행 실패")
                }

            } else {
                LogDebug("백그라운드 실행 생략: 비밀번호(webPW, pw2) 정보 부족")
            }
        } catch as e {
            LogDebug("백그라운드 실행 중 오류 발생: " e.Message)
        }

        payload := Map("type", "loginSuccess", "profile", profile)

        jsonResp := JSON.stringify(payload)
        wv.PostWebMessageAsJson(jsonResp)

        ; [추가] 근무 조 정보 전송
        userTeam := profile.Has("team") ? profile["team"] : ""
        shiftContext := WorkLogManager.GetCurrentContext(userTeam)
        shiftPayload := Map("type", "updateShiftStatus", "data", shiftContext)
        wv.PostWebMessageAsJson(JSON.stringify(shiftPayload))
        shiftPayload := Map("type", "updateShiftStatus", "data", shiftContext)
        wv.PostWebMessageAsJson(JSON.stringify(shiftPayload))

        ; [추가] 작업보고.sap 파일 갱신
        UpdateSapFile(uid)

        ; [추가] 자동 종료 타이머 시작
        StartAutoExitTimer(profile)
    }
    ; --- 2. 데이터/설정 관리 ---
    else if (command == "requestConfig") {
        cfg := ConfigManager.Config
        payload := Map("type", "loadConfig", "data", cfg)
        wv.PostWebMessageAsJson(JSON.stringify(payload))
    } else if (command == "checkSubstation") {
        location := msg.Has("location") ? msg["location"] : ""
        if (location != "") {
            ERP점검.PreCheck(location)
        }
    } else if (command == "msgbox") {
        text := msg.Has("text") ? msg["text"] : ""
        title := msg.Has("title") ? msg["title"] : "알림"
        MsgBox(text, title)
    }
    else if (command == "deleteUser") { ; 유저 삭제
        if (msg.Has("id")) {
            targetID := msg["id"]
            if (MsgBox("정말로 해당 유저를 삭제하시겠습니까?`n(ID: " targetID ")", "삭제 확인", "YesNo Icon?") == "Yes") {
                if ConfigManager.Config["users"].Has(targetID) {
                    ConfigManager.Config["users"].Delete(targetID)
                    ConfigManager.Save()
                    MsgBox("삭제되었습니다.", "알림")
                }
                profiles := ConfigManager.GetProfiles()

                payload := Map("type", "initLogin", "users", profiles)
                jsonResp := JSON.stringify(payload)
                wv.PostWebMessageAsJson(jsonResp)
            }
        }
    }
    else if (command == "addUser") { ; 유저 추가
        if (msg.Has("data")) {
            newUser := msg["data"]
            newID := newUser["id"]

            if ConfigManager.Config["users"].Has(newID) {
                payload := Map("type", "loginFail", "message", "이미 존재하는 사번입니다.")
                wv.PostWebMessageAsJson(JSON.stringify(payload))
            } else {
                ConfigManager.Config["users"][newID] := Map(
                    "profile", newUser,
                    "hotkeys", [],
                    "presets", Map()
                )
                ConfigManager.Save()

                profiles := ConfigManager.GetProfiles()
                payload := Map("type", "initLogin", "users", profiles)
                wv.PostWebMessageAsJson(JSON.stringify(payload))
            }
        }
    }
    ; --- 3. UI 제어 ---
    else if (command == "minimize") {
        MainGui.Minimize()
    }
    else if (command == "close") {
        OnWindowClose() ; 최소화 로직 재사용
    }
    else if (command == "saveConfig") { ; 설정 저장
        if msg.Has("data") {
            newConfig := msg["data"]
            ConfigManager.Config := newConfig
            ConfigManager.Save()

            ; [Hot Reload] 즉시 적용 로직 (백엔드만 갱신)
            ; UI 리로드 메시지(initLogin, loadConfig)는 입력 포커스를 뺏으므로 제거함.

            ; 1. CurrentUser 재연결 (Config 객체가 교체되었으므로 참조 갱신)
            uid := ConfigManager.GetCurrentUserID()
            if (uid != "") {
                ConfigManager.CurrentUser := ConfigManager.GetUserRoot(uid)["profile"]
            }

            ; 2. AHK 내부 갱신 (전역변수, Types, Orders, Hotkeys)
            LoadConfigData()

            ; 3. ERP 점검 상태 갱신 (설정창과 무관하므로 유지)
            ERP점검.ValidLocations := Map() ; 캐시 초기화
            ERP점검.RequestStatus() ; 상태 갱신 재요청
        }
    }
    ; --- 4. 매크로 제어 ---
    else if (command == "runTask") {
        taskName := msg.Has("task") ? msg["task"] : ""

        ; 메인 창 최소화 및 로딩 GUI 표시
        ShowLoadingGUI()

        ; 작업을 비동기식으로 실행x 스레드 충돌 허용x
        RunTaskAsync(taskName, msg)
    }
    else if (command == "stopTask") {
        StopMacro()
    }
    else if (command == "dragWindow") { ; 창 드래그 (커스텀 타이틀바용)
        DllCall("User32\ReleaseCapture")
        PostMessage(0xA1, 2, 0, , "ahk_id " MainGui.Hwnd)
    }
    else if (command == "getReleaseNotes") { ; 릴리즈 노트 조회
        repo := "MyungjinSong/TotalAutomation"

        url := "https://api.github.com/repos/" repo "/releases?per_page=100"

        try {
            whr := ComObject("WinHttp.WinHttpRequest.5.1")
            whr.Open("GET", url, true)
            whr.Option[4] := 13056
            whr.Send()
            whr.WaitForResponse()

            if (whr.Status == 200) {
                releases := JSON.parse(whr.ResponseText)
                payload := Map("type", "releaseNotes", "data", releases)
                wv.PostWebMessageAsJson(JSON.stringify(payload))
            } else {
                wv.PostWebMessageAsJson(JSON.stringify(Map("type", "releaseNotes", "error", "GitHub API Error: " whr
                    .Status
                )))
            }
        } catch as e {
            wv.PostWebMessageAsJson(JSON.stringify(Map("type", "releaseNotes", "error", e.Message)))
        }
    }

    ; --- 5. 기타 및 상태 저장 ---
    else if (command == "saveUiState") {
        if (msg.Has("data")) {
            uiState := msg["data"]
            stateFile := A_ScriptDir "\.restore_state.json"
            try {
                FileOpen(stateFile, "w", "UTF-8").Write(JSON.stringify(uiState))
            }
        }
        Reload
    }
    ; --- 6. 엑셀 승인 정보 요청 ---
    else if (command == "req_approval_info") {
        driverName := msg.Has("driverName") ? msg["driverName"] : ""
        result := bringApproval(driverName) ; 차량일지.ahk 함수 호출

        ; 결과 전송
        payload := Map("type", "res_approval_info", "data", result)
        wv.PostWebMessageAsJson(JSON.stringify(payload))
    }
    ; --- 7. ERP Webapp 및 목업 헤드리스 갱신 요청 ---
    else if (command == "refreshERPOrder") {
        LogDebug("ERP 주문 갱신 요청 수신")
        try {
            ERP점검.RequestStatus()

            ; [Headless] 실제 브라우저 연결 (Attach Mode)
            LogDebug("Headless 연결 시도 (Attach Mode)...")
            headless := HeadlessAutomation(true, LogDebug)
            orders := headless.GetOrderList(5129)
            LogDebug("ERP 주문 목록 조회 완료. 개수: " orders.Length)

            ; 결과 전송
            payload := Map("type", "updateERPOrderList", "orders", orders)
            jsonPayload := JSON.stringify(payload)
            LogDebug("ERP 결과 전송 (길이): " StrLen(jsonPayload))
            wv.PostWebMessageAsJson(jsonPayload)
        } catch as e {
            LogDebug("Headless 오류 (refreshERPOrder): " e.Message)
        }
    }
    ; --- 7-2. 작업자 명단 불러오기 (Headless) ---
    else if (command == "importWorkers") {
        LogDebug("작업자 명단 불러오기 요청 수신")
        try {
            LogDebug("Headless 연결 시도 (Attach Mode)...")
            headless := HeadlessAutomation(true, LogDebug)

            arbpl := msg.Has("arbpl") ? msg["arbpl"] : "5129"
            LogDebug("작업장 코드: " arbpl)

            workers := headless.GetWorkerList(arbpl)
            LogDebug("작업자 명단 조회 완료. 개수: " workers.Length)

            if (workers.Length > 0) {
                payload := Map("type", "updateWorkerList", "data", workers)
                wv.PostWebMessageAsJson(JSON.stringify(payload))
                MsgBox(workers.Length "명의 작업자 정보를 불러왔습니다.", "성공")
            } else {
                LogDebug("작업자 명단이 비어있음")
                MsgBox("불러올 작업자가 없거나 조회에 실패했습니다.")
            }
        } catch as e {
            LogDebug("Headless 오류 (importWorkers): " e.Message)
            MsgBox("조회 실패: " e.Message)
        }
    }
    ; --- 8. 설정창 닫힘 (UI 갱신 요청) ---
    else if (command == "exitSettings") {
        ; 1. UI 갱신: 유저 정보 및 프로필
        profiles := ConfigManager.GetProfiles()
        payload := Map("type", "initLogin", "users", profiles)
        wv.PostWebMessageAsJson(JSON.stringify(payload))

        ; 2. UI 갱신: 설정창 데이터 및 프리셋
        cfg := ConfigManager.Config
        payload := Map("type", "loadConfig", "data", cfg)
        wv.PostWebMessageAsJson(JSON.stringify(payload))

        ; 3. UI 갱신: ERP 점검 상태
        ERP점검.ValidLocations := Map() ; 캐시 초기화
        ERP점검.RequestStatus() ; 상태 갱신 재요청
    }
}

; ==============================================================================
; 단축키 로직
; ==============================================================================
LoadConfigData() {
    global types, Orders
    global ID, PW

    ; 1. 사용자 식별
    uid := ConfigManager.GetCurrentUserID()
    if (uid == "")
        return

    ; 2. 사용자 루트 객체 로드
    userRoot := ConfigManager.GetUserRoot(uid)

    ; 3. ERP 데이터(types, Orders) 로드 (ERP점검.ahk 호환성)
    types := Map()
    Orders := Map()

    appSettings := ConfigManager.Get("appSettings", Map())
    if (appSettings.Has("locations")) {
        for loc in appSettings["locations"] {
            ; location 객체 구조: {name: "변전소A", type: "변전소", order: "1234"}
            if (loc.Has("name")) {
                name := loc["name"]
                if (loc.Has("type"))
                    types[name] := loc["type"]
                if (loc.Has("order"))
                    Orders[name] := loc["order"]
            }
        }
    }

    ; 4. 단축키 설정
    ; 4. 단축키 설정
    static RegisteredHotkeys := [] ; 이전에 등록된 핫키들을 기억하는 정적 배열

    ; 기존 핫키 모두 해제
    for hkKey in RegisteredHotkeys {
        try Hotkey hkKey, "Off"
    }
    RegisteredHotkeys := [] ; 목록 초기화

    if (userRoot.Has("hotkeys")) {
        for hk in userRoot["hotkeys"] {
            ; 활성화 확인
            isEnabled := (hk.Has("enabled") && hk["enabled"])

            if (hk.Has("key") && hk["key"] != "") {

                keyName := hk["key"]

                ; 핫키 등록 시도 (활성 상태일 때만)
                if (isEnabled) {
                    try {
                        action := hk["action"]

                        if (action == "ForceExit") {
                            Hotkey keyName, ShortcutActions.ForceExitAction, "On"
                            RegisteredHotkeys.Push(keyName)
                        }
                        else if (action == "AutoLogin") {
                            Hotkey keyName, ShortcutActions.AutoLoginAction, "On"
                            RegisteredHotkeys.Push(keyName)
                        }
                        else if (action == "OpenLog") {
                            Hotkey keyName, ShortcutActions.OpenLogAction, "On"
                            RegisteredHotkeys.Push(keyName)
                        }
                        else if (action == "ConvertExcel") {
                            Hotkey keyName, ShortcutActions.ConvertExcelAction, "On"
                            RegisteredHotkeys.Push(keyName)
                        }
                        else if (action == "CopyExcel") {
                            Hotkey keyName, ShortcutActions.CopyExcelAction, "On"
                            RegisteredHotkeys.Push(keyName)
                        }
                        else if (action == "PasteExcel") {
                            Hotkey keyName, ShortcutActions.PasteExcelAction, "On"
                            RegisteredHotkeys.Push(keyName)
                        }
                    } catch as e {
                        ; 유효하지 않은 키 무시
                    }
                }
            }
        }
    }

    ; 5. 전역 변수 설정 (레거시 호환)
    if (ConfigManager.CurrentUser.Has("id"))
        ID := ConfigManager.CurrentUser["id"]
    else
        ID := ""

    if (ConfigManager.CurrentUser.Has("sapPW"))
        PW := ConfigManager.CurrentUser["sapPW"]
    else
        PW := ""
}

; ==============================================================================
; 로딩 GUI 및 매크로 제어
; ==============================================================================
ShowLoadingGUI() {
    global MainGui, LoadingGui

    MainGui.Minimize()

    ; ToolWindow 생성
    LoadingGui := Gui("+AlwaysOnTop -Caption +ToolWindow", "MacroRunning")
    LoadingGui.BackColor := "D3D3D3"

    ; 크기 계산 (200x70) - 가로 레이아웃에 더 적합
    w := 200
    h := 70
    ; 위치: 너비 80%, 높이 20%
    x := A_ScreenWidth * 0.8
    y := A_ScreenHeight * 0.2

    ; 화면 밖으로 나가지 않도록 확인
    if (x + w > A_ScreenWidth)
        x := A_ScreenWidth - w - 20

    LoadingGui.Show("x" x " y" y " w" w " h" h " NoActivate")

    ; GIF 및 텍스트용 ActiveX 생성 (Shell.Explorer)
    try {
        wb := LoadingGui.Add("ActiveX", "x0 y0 w" w " h" h, "Shell.Explorer").Value

        gifPath := A_ScriptDir "\ui\img\loading.gif"

        html :=
            "<html><body style='margin:0; padding:0; overflow:hidden; background-color:#D3D3D3; height:100%; border:none; font-family:`"Malgun Gothic`";'>"
            . "<table width='100%' height='100%' cellpadding='0' cellspacing='0' border='0'>"
            . "<tr>"
            . "<td align='right' width='60' style='padding-right:10px;'><img src='" gifPath "' width='40' height='40' style='display:block;'></td>"
            . "<td align='left' valign='middle'>"
            . "<div style='font-size:15px; font-weight:bold; color:#333; margin-bottom:2px;'>진행중</div>"
            . "<div style='font-size:11px; color:#555;'>(중지: ESC)</div>"
            . "</td>"
            . "</tr>"
            . "</table>"
            . "</body></html>"

        wb.Navigate("about:blank")
        wb.document.write(html)
        wb.silent := true
    } catch {
        LoadingGui.Add("Text", "x0 y30 w" w " Center", "Loading...")
    }

    ; ESC 단축키 등록
    Hotkey "Esc", StopMacro, "On"
}

StopMacro(*) {

    ; [중요] 실행 중인 외부 프로세스(curl) 정리
    ; 다운로드 중 재시작 시 좀비 프로세스 방지 및 파일 잠금 해제
    try {
        RunWait "taskkill /F /IM curl.exe", , "Hide"
        ; 임시 JSON 파일 정리 (ERP 점검용)
        try FileDelete(A_WorkingDir "\temp_*.json")
    }

    Hotkey "Esc", "Off"

    ; [Refactor] 즉시 Reload (상태는 RunTaskAsync 시작 시 저장됨)
    Reload
}

EndMacro(*) {
    global MainGui, LoadingGui

    ; GUI 정리
    if (LoadingGui) {
        LoadingGui.Destroy()
        LoadingGui := ""
    }

    MainGui.Show()
}

; ==============================================================================
; 비동기 작업 실행기
; ==============================================================================
; 설명:
;   WebView에서 'runTask' 명령으로 전달된 작업을 수행합니다.
; 'msg' 객체에는 UI에서 보낸 모든 파라미터(옵션, 날짜, 타겟 등)가 들어있으므로 이를 적극 활용하세요.
; ==============================================================================
RunTaskAsync(taskName, msg) {

    MainGui.Hide()

    ; --------------------------------------------------------------------------
    ; 공통 데이터 주입 (전역 변수 의존성 제거)
    ; --------------------------------------------------------------------------
    currentID := ConfigManager.GetCurrentUserID()
    if (currentID != "") {
        msg["ID"] := currentID
        userRoot := ConfigManager.GetUserRoot(currentID)
        if (userRoot.Has("profile")) {
            profile := userRoot["profile"]
            msg["department"] := profile.Has("department") ? profile["department"] : ""
            msg["webPW"] := profile.Has("webPW") ? profile["webPW"] : ""
            msg["sapPW"] := profile.Has("sapPW") ? profile["sapPW"] : ""
        } else {
            msg["webPW"] := ""
            msg["sapPW"] := ""
        }

        ; 빠른 재시작을 위해 ID 저장
        try {
            FileOpen(A_ScriptDir "\.restart_login", "w").Write(currentID)
        }

        ; 앱에서 전달받은 UI 상태 저장 (있을 경우)
        if (msg.Has("uiState")) {
            stateFile := A_ScriptDir "\.restore_state.json"
            try {
                FileOpen(stateFile, "w", "UTF-8").Write(JSON.stringify(msg["uiState"]))
            }
        }
    }

    ; 1. 업무일지
    if (taskName == "createWorkLog") {

        workLog := RunWorkLog(msg.Has("data") ? msg["data"] : map())
        if workLog.Run() {
            EndMacro()
            MsgBox("일지작성이 완료되었습니다.", "알림", "Iconi 262144")
        }

    }
    ; 2. ERP 점검 (변전소/전기실 등)
    else if (taskName == "ERPCheck") {

        ERP점검.Start(msg)

        EndMacro()
    }
    ; 3. 선로 출입 일지
    else if (taskName == "TrackAccess") {

        RunTrackAccess(msg.Has("data") ? msg["data"] : map())

        EndMacro()
    }
    ; 4. 차량 운행 일지
    else if (taskName == "VehicleLog") {

        RunVehicleLog(msg.Has("data") ? msg["data"] : map())

        EndMacro()
    }
    ; 4-1. 승인정보 가져오기
    else if (taskName == "bringApproval") {

        result := bringApproval(msg.Has("data") ? msg["data"] : map())

        if (result) {
            payload := Map("type", "approvalInfo", "data", result)
            wv.PostWebMessageAsJson(JSON.stringify(payload))
        }

        EndMacro()
    }
    ; 5. 알 수 없는 작업 처리
    else {
        MsgBox("알 수 없는 작업: " taskName)
        StopMacro()
    }
}

; ==============================================================================
; SAP 파일 관리 함수
; ==============================================================================
UpdateSapFile(userId) {
    if (userId == "")
        return

    sapPath := A_ScriptDir "\작업보고.sap"

    ; 1. 파일이 없으면 템플릿 생성
    if !FileExist(sapPath) {
        template := "[System]`n"
            . "Name=BEP`n"
            . "Description=02. ERP 운영시스템`n"
            . "Client=100`n"
            . "[User]`n"
            . "Name=`n"
            . "Language=KO`n"
            . "[Function]`n"
            . "Title=작업완료보고`n"
            . "Command=ZPMM2418`n"
            . "[Configuration]`n"
            . "GuiSize=Maximized`n"
            . "WorkDir=C:\Users\user\Documents\SAP\SAP GUI`n"
            . "[Options]`n"
            . "Reuse=1`n"

        try {
            FileAppend(template, sapPath, "UTF-8")
        } catch as e {
            MsgBox("SAP 파일 생성 실패: " e.Message)
            return
        }
    }

    ; 2. 유저 ID 업데이트 (IniWrite 사용)
    try {
        IniWrite(userId, sapPath, "User", "Name")
    } catch as e {
        ; 파일이 사용 중이거나 권한 문제 등 무시
    }
}

; ==============================================================================
; 자동 종료 기능
; ==============================================================================
global AutoExitTarget := ""

StartAutoExitTimer(profile) {
    global AutoExitTarget

    ; 1. 옵션 확인 (autoExit: true 일 때만 동작)
    if (!profile.Has("autoExit") || !profile["autoExit"]) {
        SetTimer CheckAutoExit, 0 ; 타이머 해제
        AutoExitTarget := ""
        return
    }

    ; 2. 근무 컨텍스트 확인
    userTeam := profile.Has("team") ? profile["team"] : ""
    context := WorkLogManager.GetCurrentContext(userTeam)
    isNight := context["isNight"]

    ; 3. 목표 시간 설정 (주간 18:00, 야간 09:00)
    if (isNight)
        AutoExitTarget := "0900"
    else
        AutoExitTarget := "1800"

    ; 4. 타이머 시작 (30초 간격)
    SetTimer CheckAutoExit, 30000

    ; 즉시 한 번 체크 (로그인 시점이 종료 시간일 수 있음)
    CheckAutoExit()
}

CheckAutoExit() {
    global AutoExitTarget
    if (AutoExitTarget == "")
        return

    currentHHMM := FormatTime(A_Now, "HHmm")

    ; 목표 시간과 정확히 일치하면 종료
    if (currentHHMM == AutoExitTarget) {
        OnExitApp()
    }
}

OnBackgroundLog(line) {
    if (line == "")
        return

    ; [Debug] 모든 로그 기록
    LogDebug("[BG] " line)

    ; [Cookie] 쿠키 수신
    if (SubStr(line, 1, 7) == "COOKIE:") {
        cookie := SubStr(line, 8)
        ; [Debug] 각 쿠키의 키와 값 일부(앞뒤 4자리)를 로그에 기록
        logKeys := ""
        loop parse, cookie, ";" {
            kp := Trim(A_LoopField)
            if (eqPos := InStr(kp, "=")) {
                k := SubStr(kp, 1, eqPos - 1)
                v := SubStr(kp, eqPos + 1)

                if (StrLen(v) > 8)
                    dispVal := SubStr(v, 1, 4) . ".." . SubStr(v, -4)
                else
                    dispVal := v

                logKeys .= k . ":" . dispVal . ", "
            }
        }
        LogDebug("쿠키 수신: [" . RTrim(logKeys, ", ") . "]")

        return
    }

    ; [State] 상태 메시지 처리
    if (SubStr(line, 1, 6) == "STATE:") {
        state := SubStr(line, 7)
        LogDebug("[BG State] " state)

        if (state == "BG_START") {
            ;TrayTip "백그라운드 프로세스가 시작되었습니다.", "알림", 1
        }
        else if (state == "LOGIN_OK") {
            HandleLoginSuccess()
        }
        else if (state == "LOGIN_FAIL") {
            HandleLoginFail()
        }
        return
    }

    ; [Legacy Compatibility] "Ready", "Fail" 문자열 처리
    if (line == "Ready") {
        HandleLoginSuccess()
        return
    }
    if (line == "Fail") {
        HandleLoginFail()
        return
    }

    ; [Default] 일반 메시지는 타이틀바 업데이트 (사용자 요청으로 제거)
    ; userName := (ConfigManager.CurrentUser.Has("name")) ? ConfigManager.CurrentUser["name"] : "사용자"
    ; baseTitle := "통합자동화 v3 - " . userName
    ; statusTitle := baseTitle . " - " . line

    ; if (wv) {
    ;    payload := Map("type", "updateTitle", "title", statusTitle)
    ;    wv.PostWebMessageAsJson(JSON.stringify(payload))
    ; }
}

HandleLoginSuccess() {
    global wv
    userName := (ConfigManager.CurrentUser.Has("name")) ? ConfigManager.CurrentUser["name"] : "사용자"
    baseTitle := "통합자동화 v3 - " . userName

    readyTitle := baseTitle . " - ERP 연결 중..."
    if (wv) {
        payload := Map("type", "updateTitle", "title", readyTitle)
        wv.PostWebMessageAsJson(JSON.stringify(payload))
    }

    try {
        headless := HeadlessAutomation.Connect(LogDebug)
        finalTitle := baseTitle . " - 일지연동 조회준비 완료"
        if (wv) {
            payload := Map("type", "updateTitle", "title", finalTitle)
            wv.PostWebMessageAsJson(JSON.stringify(payload))
            wv.PostWebMessageAsJson(JSON.stringify(Map("type", "headlessReady")))
        }
    } catch as e {
        LogDebug("Headless 연결 실패: " e.Message)
        errTitle := baseTitle . " - 연결 실패"
        if (wv) {
            payload := Map("type", "updateTitle", "title", errTitle)
            wv.PostWebMessageAsJson(JSON.stringify(payload))
        }
    }

    BackgroundProcessManager.Cleanup()
}

HandleLoginFail() {
    global wv
    userName := (ConfigManager.CurrentUser.Has("name")) ? ConfigManager.CurrentUser["name"] : "사용자"
    baseTitle := "통합자동화 v3 - " . userName

    failTitle := baseTitle . " - 연동조회 로그인 실패"
    if (wv) {
        payload := Map("type", "updateTitle", "title", failTitle)
        wv.PostWebMessageAsJson(JSON.stringify(payload))
    }
    MsgBox("로그인에 실패했습니다.", "실패")
    BackgroundProcessManager.Cleanup()
}

; ------------------------------------------------------------------------------
; 세션BG 출력 처리 콜백
; ------------------------------------------------------------------------------
OnSessionBGOutput(line) {
    global wv
    try {
        ; JSON 파싱 시도 (한 줄씩 들어옴)
        if (SubStr(Trim(line), 1, 1) == "{") {
            data := JSON.parse(line)

            if (data.Has("error")) {
                LogDebug("세션BG 오류: " data["error"])
                return
            }

            ; 쿠키 저장 (HeadlessAutomation 전역 저장소)
            LogDebug("세션BG: 쿠키 데이터 수신 타입: " Type(data))
            if (Type(data) == "Map")
                LogDebug("세션BG: 쿠키 데이터 키 개수: " data.Count)
            else
                LogDebug("세션BG: 쿠키 데이터 키 개수 체크 불가 (Map 아님)")

            HeadlessAutomation.SetCookieStorage(data)
            LogDebug("세션BG: 쿠키 저장 완료 (btcep, niw, ep)")

            ; 세션BG에게 쿠키 저장 완료 메시지 전송
            BackgroundProcessManager.SendInput("COOKIE_SAVED")

            ; UI에 Headless 준비 완료 알림
            if (wv) {
                wv.PostWebMessageAsJson(JSON.stringify(Map("type", "headlessReady")))
            }

        } else {
            ; 일반 로그 (필요시)
            ; LogDebug("세션BG(Log): " line)
        }
    } catch as e {
        LogDebug("세션BG 출력 파싱 실패: " e.Message " / 내용: " line)
    }
}
