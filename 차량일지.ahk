; RunVehicleLog - 차량운행일지 자동 입력 매크로
; data: Frontend에서 전달받은 JSON 객체

RunVehicleLog(data) {

    if !data
        return

    ; --- 1. 브라우저/로그인 확인 ---
    if !cUIA := WebAutoLogin.EnsureReady("WorkLog_View") {
        LogDebug("[오류] cUIA 반환 실패 (RunVehicleLog)")
        MsgBox "cUIA 반환 실패"
        return
    }
    else
        thisId := cUIA.BrowserId

    ; 메뉴 클릭
    menuClick(cUIA, "분야업무")

    ; "추가 조회 삭제" 등의 메뉴에서 "btnReg" (등록/추가) 버튼 클릭
    try {
        cUIA.WaitElement({ Name: "추가 조회 삭제" }).WaitElement({ AutomationId: "btnReg" }, 3000).Invoke()
    } catch {
        LogDebug("[오류] 차량일지 등록 버튼을 찾을 수 없음")
        MsgBox("차량일지 등록 버튼을 찾을 수 없습니다.")
        return
    }

    if !WinWaitNotActive(thisId, , 5) {
        LogDebug("[오류] 차량일지 추가 페이지 오픈 실패 Timeout")
        MsgBox("차량일지 추가 페이지 오픈 실패 Timeout")
        return
    }

    carPage := WinExist("A")
    cUIA := UIA_Browser(carPage)

    ; --- 2. 입력 로직 실행 ---

    ;차량번호 선택창
    cUIA.WaitElement({ Type: "Image", Name: "차량번호" }, 3000).Invoke()

    if !WinWaitNotActive(carPage, , 5) {
        LogDebug("[오류] 차량선택 페이지 오픈 실패 Timeout")
        MsgBox("차량선택 페이지 오픈 실패 Timeout")
        ExitApp
    }

    cUIA_Sub := UIA_Browser(WinExist("A"))

    ; 차량번호 선택
    try
        cUIA_Sub.WaitElement({ LocalizedType: "텍스트", OR: [{ Name: "타워모터카", mm: 2 }, { Name: "하이브리드모터카", mm: 2 }] },
        35000).ControlClick()

    ;사용작업장 선택창
    cUIA.WaitElement({ Type: "Image", Name: "작업장" }, 3000, , index := 2).Invoke()

    if !WinWaitNotActive(carPage, , 5) {
        LogDebug("[오류] 작업장 페이지 오픈 실패 Timeout")
        MsgBox("작업장 페이지 오픈 실패 Timeout")
        ExitApp
    }

    cUIA_Sub := UIA_Browser(WinExist("A"))

    ;사용작업장
    namebox := cUIA_sub.FindElement({ AutomationId: "I_KTEXT" })
    namebox.value := data["department"]
    namebox.SetFocus()
    Sleep 250
    cUIA_Sub.send "{enter}"
    Sleep 250
    try
        cUIA_Sub.FindElement({ LocalizedType: "텍스트", Name: data["department"] }).ControlClick()

    ;운전자
    cUIA.WaitElement({ Type: "Image", Name: "요청인" }, 3000, , index := 2).Invoke()
    lensInput(carPage, , data["driver"])

    ;작업시점 선택창
    cUIA.FindAll({ Type: "Image", Name: "위치" })[1].Invoke()

    if !WinWaitNotActive(carPage, , 5) {
        LogDebug("[오류] 작업시점 페이지 오픈 실패 Timeout (1차)")
        MsgBox("작업시점 페이지 오픈 실패 Timeout")
        ExitApp
    }

    cUIA_Sub := UIA_Browser(WinExist("A"))

    ;작업구간 시점
    namebox := cUIA_sub.FindElement({ AutomationId: "I_ZZLOCTEXT" })
    namebox.value := data["point1"]
    namebox.SetFocus()
    Sleep 250
    cUIA_Sub.send "{enter}"
    Sleep 250

    try
        cUIA_sub.WaitElement({ LocalizedType: "그룹", Name: data["point1"] }).ControlClick()

    ;작업시점 선택창
    cUIA.FindAll({ Type: "Image", Name: "위치" })[2].Invoke()

    if !WinWaitNotActive(carPage, , 5) {
        LogDebug("[오류] 작업시점 페이지 오픈 실패 Timeout (2차)")
        MsgBox("작업시점 페이지 오픈 실패 Timeout")
        ExitApp
    }

    cUIA_Sub := UIA_Browser(WinExist("A"))
    ;작업구간 종점
    namebox := cUIA_sub.FindElement({ AutomationId: "I_ZZLOCTEXT" })
    namebox.value := data["point2"]
    namebox.SetFocus()
    Sleep 250
    cUIA_Sub.send "{enter}"
    Sleep 250

    try
        cUIA_sub.WaitElement({ LocalizedType: "그룹", Name: data["point2"] }).ControlClick()

    ; 선로구분
    list_obj := cUIA.WaitElement({ AutomationId: "I_LINGB" })
    list_obj.expand()
    Sleep 50
    list_obj.waitelement({ Name: data["trackType"] }, 1000).invoke()
    Sleep 50
    list_obj.collapse()
    Sleep 100

    ; 텍스트 필드 입력
    cUIA.WaitElement({ AutomationId: "I_PERMITN" }).value := data["approveNo"]
    cUIA.WaitElement({ AutomationId: "I_JOBCONT" }).value := data["content"]
    cUIA.WaitElement({ AutomationId: "I_ZBIGO" }).value := data["remarks"]

    ; 0으로 초기화
    cUIA.WaitElement({ AutomationId: "I_CDAY_QTY_P" }).value := "0"
    cUIA.WaitElement({ AutomationId: "I_CDAY_QTY" }).value := "0"
    cUIA.WaitElement({ AutomationId: "I_CDAY_QTY_M" }).value := "0"

    workDate := FormatTime(, "yyyy-MM-dd")
    cUIA.FindElement({ AutomationId: "I_SERV_DATEF" }).value := workDate
    cUIA.FindElement({ AutomationId: "I_SERV_TIMEF" }).value := data["startTime"]
    cUIA.FindElement({ AutomationId: "I_SERV_DATET" }).value := workDate
    cUIA.FindElement({ AutomationId: "I_SERV_TIMET" }).value := data["endTime"]
    Sleep 250
    cUIA.send "{enter}"

    ; 적산계 계산
    currentSum := cUIA.FindElement({ AutomationId: "I_TT_SUM" }).value
    if (data["distance"] != "" && IsNumber(data["distance"]))
        runDist := Round(data["distance"] - currentSum)
    else
        runDist := 0

    cUIA.FindElement({ AutomationId: "I_TT_CDAY" }).value := runDist

    ; 가동시간
    cUIA.FindElement({ AutomationId: "I_OT_CDAY" }).value := data["runTime"]
    Sleep 250
    cUIA.send "{enter}"

    ; 승인자 입력 (돋보기)
    try {
        cUIA.WaitElement({ Type: "Image", Name: "승인자" }, 3000).Invoke()
        lensInput(carPage, data["dept"], data["approver"])
    }

    MsgBox "입력을 완료하였습니다. 확인 후 저장해 주세요", "통합자동화", "iconi"
    WinActivate carPage

    return
}

bringApproval(data) {
    global wv

    driverName := data.Has("driverName") ? Trim(data["driverName"]) : ""
    if (driverName = "") {
        PostApprovalError("운전자 이름이 없습니다.")
        return
    }

    if SessionManager.IsReady() {
        BringApprovalQuery(driverName, false)
        return
    }

    onReady := (*) => BringApprovalQuery(driverName, true)
    onError := (message) => PostApprovalError(message)
    SessionManager.AcquireAsync(ConfigManager.CurrentUser, onReady, onError)
}

BringApprovalQuery(driverName, retried := false) {
    global wv

    try {
        result := ApprovalInfoService.FindFirst(driverName)
        payload := Map(
            "type", "approvalInfo",
            "data", Map(
                "승인번호", result["ApprovalNo"],
                "승인부서", result["Dept"],
                "승인자", result["Approver"]
            )
        )
        wv.PostWebMessageAsJson(JSON.stringify(payload))
    } catch as err {
        if !retried && SessionManager.IsExpiredError(err) {
            LogDebug("승인정보 조회 중 세션 만료 감지: 통합 세션을 1회 재획득합니다.")
            onReady := (*) => BringApprovalQuery(driverName, true)
            onError := (message) => PostApprovalError(message)
            SessionManager.Reacquire(onReady, onError)
            return
        }
        PostApprovalError(err.Message)
    }
}

PostApprovalError(message) {
    global wv
    LogDebug("[승인정보 조회 실패] " message)
    payload := Map("type", "approvalInfo", "data", false, "error", message)
    try wv.PostWebMessageAsJson(JSON.stringify(payload))
}

; --- Helper Functions ---

lensInput(originalID, office := "", sname := "") {

    if !WinWaitNotActive(originalId, , 5) {
        LogDebug("[오류] lensInput Timeout")
        MsgBox("ERROR - lensInput Timeout", "timeout", "icon!")
        return
    }

    cUIA_sub := UIA_Browser(WinExist("A"))
    officebox := cUIA_sub.WaitElement({ AutomationId: "I_STEXT" }, 3000)
    namebox := cUIA_sub.FindElement({ AutomationId: "I_SNAME" })

    if office {
        officebox.value := office
        namebox.value := sname
        namebox.SetFocus()
        Sleep 250
        cUIA_sub.send "{enter}"
    }
    else {
        ; 부서 없이 이름만 검색하는 경우
        namebox.value := sname
        namebox.SetFocus()
        Sleep 250
        cUIA_sub.send "{enter}"
    }

    ; 검색 결과가 뜰 때까지 대기 후 클릭
    cUIA_sub.WaitElement({ LocalizedType: "그룹", Name: sname }, 3000).ScrollintoView()
    Sleep 500

    try
        cUIA_sub.WaitElement({ LocalizedType: "그룹", Name: sname }, 3000).ControlClick()

    sleep 250
}

menuClick(cUIA, str) {
    try {
        ; 상단 메뉴 바 찾기
        menubtn := cUIA.WaitElement({ Name: "인원현황 일반업무 주요업무 자재사용 분야업무 안전관리 운전적합성 점검표" }, 5000)
        Sleep 100
        menubtn.FindElement({ Name: str }).ControlClick()
    } catch {
        LogDebug("[오류] 메뉴 클릭 실패: " str)
        MsgBox("메뉴 클릭 실패: " str)
    }
    return
}

WinWaitNotActive(winTitle, winText := "", timeout := 0) {
    startTime := A_TickCount
    while true {
        if !WinActive(winTitle, winText) {
            return true
        }
        if (timeout > 0 && (A_TickCount - startTime) / 1000 >= timeout) {
            return false
        }
        Sleep(100)
    }
}
