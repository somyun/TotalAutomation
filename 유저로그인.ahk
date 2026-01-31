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

    ; ==============================================================================
    ; [메서드] EnsureReady
    ; 설명: 작업 유형에 따른 브라우저 상태를 준비합니다.
    ; 매개변수:
    ;   taskType - 작업 유형 ("WorkLog_Create", "WorkLog_View", "SessionCheck")
    ; 반환값: 성공 시 해당 브라우저 UIA_browser 객체, 실패 시 false
    ; ==============================================================================
    static EnsureReady(taskType) {
        user := ConfigManager.CurrentUser
        if (!user.Has("id")) {
            MsgBox("로그인된 사용자가 없습니다. 먼저 로컬 로그인을 수행해주세요.", "오류", "Iconx")
            return false
        }

        ; 1. 브라우저/세션 점검 및 로그인
        if (taskType == "SessionCheck") {
            ; 세션 브라우저 cUIA 반환
            return this._PrepareSessionOnly(user)
        }
        else if (taskType == "WorkLog_Create") {
            ; 세션 브라우저 cUIA 반환
            cUIA := this._PrepareSessionOnly(user)
            ; 해당 브라우저로 업무일지 리스트 이동
            this._GoToWorklogList(cUIA)
            ; 업무일지 생성 클릭
            cUIA.WaitElement({ LocalizedType: "링크", Name: "생성" }, 5000).Invoke()
            ; 생성된 업무일지 cUIA할당, 로딩까지 시간이 걸릴 수 있으므로 5초동안 시도
            loop 20 {
                if cUIA := this._FindBrowserByElement(create := true)
                    break
                Sleep 250
            }
            ; cUIA 반환
            return cUIA
        }
        else if (taskType == "WorkLog_View") {
            ; 1. 열려있는 업무일지 브라우저 탐색
            if cUIA := this._FindBrowserByElement()
                return cUIA
            ; 2. 없으면 ERP 포털 또는 통합포털이 열린 브라우저 탐색 (배열 순차 탐색)
            if cUIA := this._FindBrowserByTab(["ERP포털시스템 - 부산교통공사", ":: 부산교통공사 ::"])
                return this._NavToWorkLogView(cUIA, user)
            ; 3. ERP 포털도 없으면 새로 실행
            cUIA := this._PrepareSessionOnly(user)
            return this._NavToWorkLogView(cUIA, user)
        }

        return true
    }

    ; ==============================================================================
    ; [내부] _FindBrowserByElement
    ; 설명: 실행 중인 브라우저를 순회하며 특정 요소가 있는 브라우저 객체를 반환
    ; ==============================================================================
    static _FindBrowserByElement(create := false) {
        targetBrowsers := ["msedge.exe", "chrome.exe", "whale.exe"]

        for exe in targetBrowsers {
            if !ProcessExist(exe)
                continue

            if !hwndList := WinGetList("ahk_exe " exe)
                continue

            for hwnd in hwndList {
                ; [최적화] 타이틀에 "부산교통공사"가 없으면 스킵 (가장 확실한 필터)
                if !InStr(WinGetTitle(hwnd), "부산교통공사")
                    continue

                try {
                    cUIA := UIA_Browser("ahk_id " hwnd)
                    nowDate := FormatTime(DateAdd(A_Now, create ? 0 : -9, "Hours"), "yyyy-MM-dd")
                    ; UIA_Browser 인스턴스 생성만으로도 시간이 소요되므로 위 조건으로 최대한 필터링함
                    if cUIA.FindElement({ AutomationId: "I_GIJUND", Value: nowDate })
                        return cUIA
                }
            }
        }
        return false
    }

    ; ==============================================================================
    ; [내부] _FindBrowserByTab(
    ; 설명: 실행 중인 브라우저를 순회하며 탭 이름에 키워드(배열)가 포함된 브라우저 객체를 반환
    ; ==============================================================================
    static _FindBrowserByTab(keywordArr) {
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
                cUIA := UIA_Browser("ahk_id " hwnd)
                tabs := cUIA.GetAllTabNames()

                for keyword in keywordArr {
                    for tabItem in tabs {
                        if InStr(tabItem, keyword) {
                            try {
                                cUIA.SelectTab(tabItem) ; 해당 탭 선택
                            }
                            return cUIA
                        }
                    }
                }
            }
        }
        return false
    }

    ; ==============================================================================
    ; [내부] _PrepareSessionOnly (Pattern 선로출입관리)
    ; ==============================================================================
    static _PrepareSessionOnly(user) {

        exeName := this.GetBrowserExe()

        ; 1. 무조건 새 창으로 포털 접속 (기존 작업 방해 방지)
        ; --new-window: 새 창 강제
        ; --start-maximized를 빼고 실행해야 이동이 수월함
        Run(exeName " --new-window " this.PortalURL) ;PortalURL
        WinWaitActive("ahk_exe " exeName, , 5)
        WinRestore(WinExist())

        ; 2. UIA 연결
        try {
            ; WinWait로 찾은 마지막 창(Last Found Window)을 대상으로 UIA 초기화
            cUIA := UIA_Browser("ahk_id " WinExist())
        } catch as e {
            ; 연결 실패 시 로그 남기거나 false 반환
            MsgBox "cUIA 연결 실패"  ;디버깅용
            return false
        }

        ; 5. 로그인 점검
        if !this.IsLoggedIn(cUIA)
            return this.Login(user["id"], user["webPW"], user["pw2"], cUIA)

        return cUIA
    }

    ; ==============================================================================
    ; [내부] _NavToWorkLogView (Pattern 일지조회)
    ; ==============================================================================
    static _NavToWorkLogView(cUIA, user) {
        try {
            ; 1. 메뉴 이동
            this._GoToWorklogList(cUIA)

            ; 2. 일지 검색 및 클릭
            targetDate := FormatTime(DateAdd(A_Now, -9, "Hours"), "yyyyMMdd")
            dept := user.Has("department") ? user["department"] : "호포전기분소"
            targetName := targetDate " " dept " 업무일지"

            try {
                ; 항목 클릭
                cUIA.WaitElement({ LocalizedType: "텍스트", Name: targetName }, 7000).Click("Left")
                Sleep 250
                ; 변경/조회 클릭
                cUIA.WaitElement({ LocalizedType: "링크", Name: "변경/조회" }, 3000).Invoke()
                Sleep 500
            } catch {
                MsgBox("오늘자 업무일지(" targetName ")를 찾을 수 없습니다.", "알림", "Icon!")
                return false
            }

            return this._FindBrowserByElement()

        } catch as e {
            MsgBox("업무일지 조회 화면 이동 중 오류: " e.Message, "오류", "Iconx")
            return false
        }
    }

    ; ==============================================================================
    ; [헬퍼] 업무일지 리스트 메뉴 이동
    ; ==============================================================================
    static _GoToWorklogList(cUIA) {

        ; 1. 현재 페이지 확인 (ERP포털이나 통합포털이 아니면 이동 필요)
        currentTitle := WinGetTitle(cUIA.BrowserId)
        isERP := InStr(currentTitle, "ERP포털시스템 - 부산교통공사")
        isIntegrated := InStr(currentTitle, ":: 부산교통공사 ::")

        if (!isERP) {
            if (isIntegrated) {
                ; 통합포털이면 ERP 버튼 클릭 시도
                try {
                    cUIA.WaitElement({ Type: "Link", Name: "ERP" }, 3000).Invoke()
                } catch {
                    ; 버튼 없으면 그냥 URL 이동
                    cUIA.navigate(this.ERP_PortalURL, , 10000)
                }
            } else {
                ; 그 외 페이지면 URL 직접 이동
                cUIA.navigate(this.ERP_PortalURL, , 10000)
            }

            ; 2. [핵심] 이동 후 상태 검증 (Smart Wait)
            ; 성공 신호(ERP 메뉴) 또는 실패 신호(로그인창 - userId) 중 먼저 뜨는 것을 10초간 대기
            try {
                foundEl := cUIA.WaitElement([{ AutomationId: "userId" }, { Name: "업무일지 업무일지" }], 10000)

                ; 3. 로그인 화면으로 튕겼는지 확인
                if (foundEl.AutomationId == "userId") {
                    ; 세션 만료됨! 재로그인 시도
                    user := ConfigManager.CurrentUser
                    if this.Login(user["id"], user["webPW"], user["pw2"], cUIA) {
                        ; 재로그인 성공 -> 다시 이동
                        cUIA.navigate(this.ERP_PortalURL, , 10000)
                        ; 이번엔 무조건 ERP 메뉴 대기 (또 실패하면 에러처리)
                        cUIA.WaitElement({ Name: "업무일지 업무일지" }, 10000)
                    } else {
                        throw Error("세션 만료 후 재로그인 실패")
                    }
                }
            } catch as e {
                MsgBox("페이지 이동 중 문제 발생: " e.Message)
                return false
            }
        }

        ; 업무일지 아이콘/버튼이 나오면 클릭 => 업무일지 리스트 페이지로 이동
        try {
            cUIA.WaitElement({ Name: "업무일지 업무일지" }, 10000).Invoke()
        } catch {
            ; 이미 리스트 등 다른 화면일 수 있으므로 패스하거나 재시도 등 고민
            ; 여기서는 일단 진행
        }

        return true
    }

    ; ==============================================================================
    ; [헬퍼] '부산교통공사'탭이 열린 브라우저 (기본:엣지)
    ; ==============================================================================
    static GetBrowserExe() {

        ; 검사할 브라우저 목록
        targetBrowsers := ["msedge.exe", "chrome.exe", "whale.exe"]

        for exe in targetBrowsers {
            ; 1. 해당 브라우저 프로세스가 없으면 스킵
            if !ProcessExist(exe)
                continue

            ; 2. 해당 브라우저의 모든 창 ID(HWND) 가져오기
            try {
                hwndList := WinGetList("ahk_exe " exe)
            } catch {
                continue
            }

            ; 3. 각 창을 순회하며 탭 검사
            for hwnd in hwndList {
                ; [최적화]
                if !InStr(WinGetTitle(hwnd), "부산교통공사")
                    continue

                try {
                    ; 최소화된 창은 UIA가 요소를 못 읽을 수 있어서 건너뛰거나, WinRestore를 해야 함.
                    ; (일단은 '조용한 탐색'을 위해 복원 없이 시도 에러나면 넘어감)

                    cUIA := UIA_Browser("ahk_id " hwnd)
                    tabs := cUIA.GetAllTabs()

                    for tabItem in tabs {
                        ; 탭 이름에 검색어가 포함되어 있는지 확인
                        if InStr(tabItem.Name, "부산교통공사") {
                            return exe ; 발견 즉시 exe 이름 반환 (함수 종료)
                        }
                    }
                } catch {
                    ; UIA 연결 실패, 권한 문제, 또는 요소 찾기 실패 시 다음 창으로 넘어감
                    continue
                }
            }
        }

        return "msedge.exe" ; 모든 브라우저를 다 뒤져도 없으면 엣지로 반환
    }

    ; ==============================================================================
    ; [메서드] IsLoggedIn
    ; 설명: Positive Validation 방식의 로그인 점검
    ; ==============================================================================
    static IsLoggedIn(cUIA := "") {
        try {
            if !cUIA
                cUIA := UIA_Browser("A") ; 현재 활성 브라우저

            loop 70 {
                ; 1. 로그아웃 버튼이 있으면 로그인된 상태 (Positive)
                try
                    if cUIA.FindElement({ Name: "로그아웃" })
                        return cUIA

                ; 2. 업무일지 메뉴 버튼이 있어도 로그인된 상태
                try
                    if cUIA.FindElement({ Name: "업무일지 업무일지" })
                        return cUIA

                ; 3. 로그인 입력창(userId)이 있으면 로그아웃 상태 (Negative)
                try
                    if cUIA.FindElement({ AutomationId: "userId" })
                        return false

                Sleep 100   ;30 * 100 = 3초간 확인
            }

            ; 4. 둘 다 없으면? 로딩중이거나 엉뚱한 페이지.
            ; 일단 False 반환하여 로그인 시도 유도하거나 예외 처리
            MsgBox "로그인 상태 확인 불가"
            return false

        } catch {
            MsgBox "알수 없는 에러"
            return false
        }
    }

    ; ==============================================================================
    ; [메서드] Login
    ; ==============================================================================
    static Login(id, pw, pw2, cUIA := "") {
        if (id == "" || pw == "") {
            MsgBox("로그인 정보(사번, 비번)가 없습니다.", "오류")
            return false
        }

        try {
            if !cUIA
                cUIA := UIA_Browser("A")

            ; 아이디 입력
            cUIA.WaitElement({ AutomationId: "userId" }, 2000).Value := id

            ; 비밀번호 입력
            cUIA.FindElement({ AutomationId: "password" }).Value := pw

            ; 로그인 버튼 클릭 (ClassName: btn_login)
            cUIA.FindElement({ ClassName: "btn_login" }).Invoke()

            ; 2차 인증 대기 (있을 경우)
            try {
                ; 인증번호 입력창 대기 (짧게)
                cUIA.WaitElement({ AutomationId: "certi_num" }, 2000).Value := pw2
                cUIA.FindElement({ ClassName: "btn_blue" }).Invoke() ; 확인 버튼
            } catch as e {
                MsgBox("2차인증 실패`n" e.Message, "오류")
                return false
            }

            ; 로그인 완료 대기 (로그아웃 버튼 뜰 때까지)
            try {
                cUIA.WaitElement({ Name: "로그아웃" }, 10000)
                Sleep 1000
                return cUIA
            } catch {
                MsgBox("로그인 후 응답이 지연되거나 실패했습니다.", "오류")
                return false
            }

        } catch as e {
            MsgBox("자동 로그인 실패: " e.Message, "오류", "Iconx")
            return cUIA
        }
    }
}
