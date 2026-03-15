; RunTrackAccess - 선로출입관리 자동 입력 매크로
; data: Frontend에서 전달받은 JSON 객체
RunTrackAccess(data) {
    if !data
        return

    ; --- 1. 브라우저/로그인 확인 ---
    if !WebAutoLogin.EnsureReady("SessionCheck") {
        MsgBox "XPLATFORM 세션 연결에 실패했습니다."
        StopMacro()
        return
    }

    if WinWait("개별업무통합관리", , 15) {
        ; 이미 켜져있는 경우, 특정 픽셀(파란색 배경 등)을 확인하여 로그인 화면이면 재로그인 시도 루틴

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
        MsgBox("세션 준비된 브라우저가 없습니다.")
        return false
    }

    if !WinWait("개별업무통합관리 - 선로출입현황 조회", , 30) {
        MsgBox "timeout - 선로출입현황 조회 화면이 뜨지 않습니다."
        return
    }

    Sleep 100
    click 110, 250   ; 선로출입 등록/신청 메뉴 클릭 (좌측 트리)
    if !WinWait("개별업무통합관리 - 선로출입 등록/신청", , 5) {
        MsgBox "timeout - 등록화면 진입 실패"
        return
    }
    Sleep 250

    originalId := WinExist("개별업무통합관리")
    WinActivate("개별업무통합관리")
    WinMaximize(originalId)

    ; --- 2. 입력 로직 실행 ---

    ; "외 N명" 문자열 생성
    othersStr := ""
    if (data.Has("totalCount") && data["totalCount"] != "") {
        deduct := 1 ; 기본 1명 (본인/대표)
        if (data["driverName"] != "")
            deduct++
        if (data["workerName"] != "")
            deduct++
        if (data["safetyName"] != "")
            deduct++

        othersCount := Integer(data["totalCount"]) - deduct
        if (othersCount > 0)
            othersStr := " 외 " othersCount "명"
    }

    ;모터카지정 여부
    모터카 := data["workType"] == 2 or data["workType"] == 4 or data["workType"] == 6

    Sleep 100

    ; (1) 행추가 (Tab 16번 -> Enter)
    Send "{tab 16}{enter}"
    Sleep 250

    ; (2) 작업구분
    Send "{tab 13}"
    Sleep 100
    try {
        loopCount := Integer(data["workType"])
        if (loopCount > 0)
            SendSleep("{down " loopCount "}", 250)
    } catch as e {
        MsgBox("작업구분 데이터 오류", "선로출입관리", "icon!")
        return
    }

    ; (3) 작업일자
    SendSleep("{tab}", 250)
    if (A_Hour < 17)
        SendSleep(FormatTime(, "yyyyMMdd"), 100)

    ; (4) 작업내역
    SendSleep("{tab}", 250)
    SendSleep(data["content"], 500)

    ; (5) 운행시작
    SendSleep("{tab}", 250)
    if (data["opStart"] != "")
        SendSleep(data["opStart"], 250)
    else
        SendSleep("^a{delete}", 250)

    ; (6) 작업시작
    SendSleep("{tab}", 250)
    SendSleep(data["workStart"], 250)

    ; (7) 운행종료
    SendSleep("{tab}", 250)
    if (data["opEnd"] != "")
        SendSleep(data["opEnd"], 250)
    else
        SendSleep("^a{delete}", 250)

    ; (8) 작업종료
    SendSleep("{tab}", 250)
    SendSleep(data["workEnd"], 250)

    ; (9) 요청사항
    SendSleep("{tab}", 100)

    ; (10) 호선
    SendSleep("{tab}", 100)
    try
        loop Integer(data["line"])
            SendSleep("{down}", 250)
    catch {
        MsgBox("호선 데이터 오류")
        return
    }

    ; (11) 선로구분
    SendSleep("{tab}", 100)
    try
        loop Integer(data["trackType"])
            SendSleep("{down}", 100)
    catch {
        MsgBox("선로구분 데이터 오류")
        return
    }

    ; (12) 선로차단여부
    SendSleep("{tab}", 100)
    Send "{down}"

    ; (13) 운행from
    SendSleep("{tab}", 100)
    if data["workType"] == 2
        Send data["workFrom"] "{enter}"

    ; (14) 작업from
    SendSleep("{tab}", 100)
    if (data["workFrom"] != "")
        Send data["workFrom"] "{enter}"

    ; (15) 경유(통과)
    SendSleep("{tab 2}", 100)

    ; (16) 작업to
    SendSleep("{tab}", 100)
    if (data["workTo"] != "")
        SendSleep(data["workTo"] "{enter}", 100)

    if 모터카 { ;모터카 출고시
        SendSleep("{tab 3}", 100)
        ;if data["bunso"] == "호포전기분소"
        SendSleep("{down}", 100)

        ;철도장비사용수량
        SendSleep("{tab}", 100)
        Send "1"
    }
    else
        SendSleep("{tab 2}", 100)

    ; (18) 운전원
    SendSleep("{tab}", 100)
    if (data["driverName"] != "")
        Send data["driverName"] "{tab}" data["driverPhone"]
    else
        SendSleep("{tab}", 100)

    ; (19) 작업자
    SendSleep("{tab}", 100)
    if (data["workerName"] != "")
        Send data["workerName"] othersStr "{tab}" data["workerPhone"]
    else
        SendSleep("{tab}", 100)

    ; (20) 총원
    SendSleep("{tab}", 100)
    Send data["totalCount"]

    ; (21) 감독자
    SendSleep("{tab 3}{enter}", 100)
    if !WinWait("직원조회", , 3) {
        MsgBox "감독자 검색 창 Timeout"
        return
    }
    Sleep 500
    if (data["supervisorId"] != "") {
        Send "+{tab}" data["supervisorId"] "{enter}"
        Sleep 750
        SendEvent "+{tab 2}"
        Sleep 100
        SendSleep("{enter}", 250)
    }
    WinWaitActive("감독자 입력", , 2)
    if WinActive("감독자 입력") {
        Sleep 100
        SendSleep("{enter}", 100)
        WinWaitClose("감독자 입력", , 3)
    }

    ; (22) 철도운행안전관리자
    SendSleep("{tab 2}", 100)
    if 모터카 ;모터카 출고시
        SendSleep("{tab 2}", 100)

    if (data["safetyName"] != "")
        Send data["safetyName"] "{tab}" data["safetyPhone"]
    else
        Send "{tab}"

    ; (23) 철도운행협의서
    SendSleep("{tab 5}", 100)
    if 모터카 { ;모터카 출고시
        SendSleep("{enter}", 100)
        if !hwnd := WinWaitActive("철도운행안전협의", , 3)
            MsgBox "철도운행안전협의 선택창 오류"

        MsgBox "철도운행안전협의서를 선택해 주세요", "선로출입관리", "iconi"
        /* 최적화된 매크로 로직 구현 후 주석 해제
        Sleep 100
        SendSleep("{tab 3}", 100)
        SendSleep("{down}", 100)
        SendSleep("{tab}", 100)
        SendSleep("{enter}", 100)
        */
        WinWaitClose(hwnd)
    }

    ; (24) 운행to
    SendSleep("{tab 12}", 100)
    if data["workType"] == 2
        Send data["workTo"] "{enter}"

    ; (25) 저장
    SendSleep("{tab 4}", 100)
    Send "{enter}"

    ; 저장 확인 및 완료 처리
    WinWaitNotActiveTime(originalId, 250)
    targetId := WinExist("A")
    Sleep 100
    Send "{enter}"

    WinWaitNotActive targetId

    WinWaitNotActiveTime(originalId, 250)
    targetId := WinExist("A")
    Sleep 100
    Send "{enter}"

    WinWaitActiveTime(originalId, 250)
    Sleep 100

    ; (26) 출입역 입력
    if (data["stationInput"]) {
        Sleep 500
        Click 300, 530
        Sleep 250
        Click 685, 540
        Sleep 250
        SendSleep("{Tab 3}" data["workFrom"], 100)
        SendSleep("{enter}", 100)
        SendSleep("{Tab}{delete}" data["totalCount"], 100)
        Send "{Tab 4}{enter}"

        if WinWaitNotActive(originalId, , 3)
            Sleep 250
        SendSleep("{enter}", 250)
        if WinWaitNotActive(originalId, , 3)
            Sleep 250
        SendSleep("{enter}", 250)

        SendEvent "+{tab 3}"
        Sleep 100
        SendSleep("{enter}", 100)
        SendSleep("{Tab 3}" data["workTo"], 100)
        SendSleep("{enter}", 100)
        SendSleep("{Tab}{delete}" data["totalCount"], 100)
        Send "{Tab 2}{enter}"

        if WinWaitNotActive(originalId, , 3)
            Sleep 2500
        SendSleep("{enter}", 250)
        if WinWaitNotActive(originalId, , 3)
            Sleep 250
        Send "{enter}"
    }

    MsgBox "입력이 완료되었습니다. 내용을 확인하세요.", "완료", "iconi"
    WinRestore(originalId)
}

; --- Helper Functions ---

SendSleep(keys, delay) {
    Send keys
    Sleep delay
}

WinWaitNotActiveTime(winTitle, millisecondsToWait, checkInterval := 50) {
    startTime := A_TickCount
    while true {
        if WinActive(winTitle) {
            startTime := A_TickCount
        } else {
            if (A_TickCount - startTime >= millisecondsToWait) {
                break
            }
        }
        Sleep(checkInterval)
    }
}

WinWaitActiveTime(winTitle, millisecondsToWait, checkInterval := 50) {
    startTime := A_TickCount
    while true {
        if WinActive(winTitle) {
            if (A_TickCount - startTime >= millisecondsToWait) {
                break
            }
        } else {
            startTime := A_TickCount
        }
        Sleep(checkInterval)
    }
}
