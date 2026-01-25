; RunVehicleLog - 차량운행일지 자동 입력 매크로
; data: Frontend에서 전달받은 JSON 객체

RunVehicleLog(data) {

    if !data
        return

    ; --- 1. 브라우저/로그인 확인 ---
    if !cUIA := WebAutoLogin.EnsureReady("WorkLog_View") {
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
        MsgBox("차량일지 등록 버튼을 찾을 수 없습니다.")
        return
    }

    if !WinWaitNotActive(thisId, , 5) {
        MsgBox("차량일지 추가 페이지 오픈 실패 Timeout")
        return
    }

    carPage := WinExist("A")
    cUIA := UIA_Browser(carPage)

    ; --- 2. 입력 로직 실행 ---

    ;차량번호 선택창
    cUIA.WaitElement({ Type: "Image", Name: "차량번호" }, 3000).Invoke()

    if !WinWaitNotActive(carPage, , 5) {
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
    cUIA_Sub.FindElement({ LocalizedType: "텍스트", Name: data["department"] }).ControlClick()

    ;운전자
    cUIA.WaitElement({ Type: "Image", Name: "요청인" }, 3000, , index := 2).Invoke()
    lensInput(carPage, , data["driver"])

    ;작업시점 선택창
    cUIA.FindAll({ Type: "Image", Name: "위치" })[1].Invoke()

    if !WinWaitNotActive(carPage, , 5) {
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

    own_id := WinExist("A")

    ; --- 1. 브라우저/로그인 확인 ---
    if !cUIA := WebAutoLogin.EnsureReady("SessionCheck") {
        MsgBox "cUIA 반환 실패"
        StopMacro()
        return
    }

    if (cUIA != "") {
        ; SSO 페이지 호출 (기존 로직 참조)

        cUIA.Navigate(
            "https://mis.humetro.busan.kr/FS/xui/install/x_installChromeSSO.jsp?gv_selSystGubn=LA&gv_userBrowser=Edg", ,
            1000)

        if WinWait("개별업무통합관리", , 15) {
            ; 이미 켜져있는 경우, 특정 픽셀(파란색 배경 등)을 확인하여 로그인 화면이면 재로그인 시도 루틴
            ; (픽셀 체크 로직은 해상도/배율에 따라 불안정할 수 있으므로, 타이틀 위주로 체크 권장)

            while !WinExist("개별업무통합관리 - 선로출입현황 조회") {

                ; 로그인 여부 확인 (450,470 좌표 색상 체크)
                if (PixelGetColor(450, 470) == 0x0063B5) {
                    targetID := WinExist("A")
                    WinClose("ahk_id " targetId)
                    MsgBox "로그인 실패"
                    return false
                }

                ; 무한 루프 방지
                if (A_Index > 5)
                    break
                Sleep 500
            }
        } else {
            ; 창이 안 뜨면 실패
            return false
        }
    } else {
        MsgBox("세션 준비된 브라우저가 없습니다.")
        return
    }

    if !WinWait("개별업무통합관리 - 선로출입현황 조회", , 30) {
        MsgBox "timeout - 선로출입현황 조회 화면이 뜨지 않습니다."
        return
    }

    WinClose("ahk_id " hwnd := cUIA.BrowserId)
    WinWaitClose("ahk_id " hwnd, , 1)

    resultData := false

    if xldata := 승인정보_엑셀추출(data["driverName"]) {
        try {
            승인번호 := xldata["승인번호"]
            승인부서 := xldata["승인부서"]
            승인자 := xldata["승인자"]
            xldata["워크북"].Close(false)  ; 저장하지 않고 닫기
            xldata["통합문서"].quit()
            resultData := Map("승인번호", 승인번호, "승인부서", 승인부서, "승인자", 승인자)
        }
        catch
            MsgBox "엑셀 자동 추출이 불가능한 상태입니다`n확인 후 직접 입력바랍니다", , "icon!"
    }
    else
        MsgBox "엑셀 자동 추출이 불가능한 상태입니다`n확인 후 직접 입력바랍니다", , "icon!"

    WinClose("개별업무통합관리")
    WinActivate(own_id)

    return resultData

}

; 승인정보_엑셀추출 - 엑셀에서 승인번호 가져오기/filtering
; driverName: 필터링할 운전원 이름
승인정보_엑셀추출(운전원) {
    if !운전원
        return { error: "운전원 이름이 없습니다." }

    ; 엑셀 실행 확인 및 연결
    try {
        try {
            xl := ComObjActive("Excel.Application")  	; 실행 중인 엑셀 붙기
            if xl.WorkBooks.Count == 0 {
                xl.quit()
                for proc in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process where Name = 'EXCEL.EXE'") {
                    proc.Terminate()
                }
            }
            else {
                for xb in xl.WorkBooks {
                    if xb.Name
                        break
                    else
                        xb.close(false)
                }
                if xl.WorkBooks.Count == 0
                    xl.quit()
            }
        }

        Sleep 1000
        Click 1200, 190									;엑셀열기
        if !WinWaitActive("Microsoft Excel - Sheet", , 10) {
            MsgBox "timeout"
            ExitApp
        }
        sleep 1000

        xl := ComObjActive("Excel.Application")  	; 실행 중인 엑셀 붙기
        wb := xl.WorkBooks("Sheet1")
        ws := wb.Worksheets("선로출입현황")
    }
    catch {
        MsgBox "엑셀 자동 추출이 불가능한 상태입니다.`n확인 후 직접 입력바랍니다", , "icon!"

        return false
    }

    try {
        ; 필터링 로직
        ; 1. 조건 설정 (A9:B10) - 위치는 적절한 빈 곳 사용
        ws.Range("A9").Value := "작업구분"
        ws.Range("B9").Value := "운전원"
        ws.Range("A10").Value := "*철도장비운행"
        ws.Range("B10").Value := 운전원

        criteriaRange := ws.Range("A9:B10")

        ; 2. 결과 복사 위치 (A11)
        ws.Range("A11:AZ100").ClearContents()
        copyToRange := ws.Range("A11")

        ; 3. 데이터 범위
        dataRange := ws.UsedRange

        ; 4. 필터 실행
        dataRange.AdvancedFilter(2, criteriaRange, copyToRange, false)

        ; 5. 결과 추출 (12행에 결과가 나온다고 가정)
        resultVal := ws.Range("A12").Value
        if (resultVal == "") {
            return { error: "해당 운전원의 승인 정보를 찾을 수 없습니다." }
        }

        ; 추출 (레거시 매핑: 승인번호=D12, 부서=AE12, 승인자=AG12)
        승인번호 := StrReplace(String(ws.Range("D12").Value), "`r`n", "")
        승인부서 := StrReplace(StrReplace(String(ws.Range("AE12").Value), " ", ""), "`r`n", "")
        승인자 := StrReplace(StrReplace(String(ws.Range("AG12").Value), " ", ""), "`r`n", "")

        return Map("승인번호", 승인번호, "승인부서", 승인부서, "승인자", 승인자, "워크북", wb, "통합문서", xl)

    } catch as e {
        return false
    }
}

; --- Helper Functions ---

lensInput(originalID, office := "", sname := "") {

    if !WinWaitNotActive(originalId, , 5) {
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
        menubtn := cUIA.WaitElement({ Name: "인원현황 일반업무 주요업무 자재사용 분야업무 안전관리 운전적합성 점검표"}, 5000)
        Sleep 100
        menubtn.FindElement({ Name: str }).ControlClick()
    } catch {
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
