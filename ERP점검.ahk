; ==============================================================================
; ERP 점검 자동화 로직
; ==============================================================================
class ERP점검 {

    ; --------------------------------------------------------------------------
    ; Entry Point
    ; --------------------------------------------------------------------------
    static ValidLocations := Map()

    ; --------------------------------------------------------------------------
    ; Entry Point
    ; --------------------------------------------------------------------------
    static Start(msg, batchmode) {
        userID := msg.Has("ID") ? msg["ID"] : ""
        userPW := msg.Has("sapPW") ? msg["sapPW"] : ""

        if (userPW == "") {
            MsgBox("SAP PW 지정되지 않아 실행할 수 없습니다", "오류", "iconx")
            return false
        }

        if (batchmode) {
            locations := msg["location"] ; Array of location objects
            downQueue := []
            locationMsg := ""

            ; 1. 다운로드 대상 수집
            for item in locations {
                loc := item["location"]
                if (item["targetType"] == "변전소" && this.ValidLocations.Has(loc)) {
                    downQueue.Push(loc)
                } else if (item["targetType"] == "변전소" && !this.ValidLocations.Has(loc)) {
                    MsgBox("웹앱에 " loc "의 데이터가 저장되어있지 않습니다.`n ERP 작업보고 일괄모드를 중단합니다.", "진행불가", "0x1 Icon!")
                    return false
                }
            }

            ; 2. 다운로드 비동기 큐 실행 (cmd & 체인)
            if (downQueue.Length > 0) {
                this.BatchDownSheetAsync(downQueue)
            }

            ; 3. 매크로 큐 실행
            for item in locations {
                loc := item["location"]
                mems := item["members"]
                type := item["targetType"]
                order := item["targetOrder"]

                memberStr := this.GetMemberStr(mems, msg.Has("format") ? msg["format"] : "summary")
                this.Macro(memberStr, loc, userID, userPW, type, order, true)
                locationMsg .= loc ", "
            }

            locationMsg := RTrim(locationMsg, ", ")
            MsgBox("ERP 일괄 입력이 완료되었습니다.`n[진행 장소: " locationMsg "]", "완료", "iconi")
            return true
        }
        else {
            ; 개별 모드
            loc := msg.Has("location") ? msg["location"] : ""
            mems := msg.Has("members") ? msg["members"] : []
            type := msg.Has("targetType") ? msg["targetType"] : ""
            order := msg.Has("targetOrder") ? msg["targetOrder"] : ""

            if (loc == "") {
                MsgBox("예외 발생 : 장소 미지정", "오류", "iconx")
                return false
            }
            if (mems.Length == 0) {
                MsgBox("예외 발생 : 작업자 미지정", "오류", "iconx")
                return false
            }

            memberStr := this.GetMemberStr(mems, msg.Has("format") ? msg["format"] : "summary")

            if (type == "변전소") {
                if (this.ValidLocations.Has(loc)) {
                    if (MsgBox("ERP 작업보고를 시작합니다.`n`n장소 : " . loc . " (WEB앱 연동)`n점검자 : " . memberStr, "진행합니다",
                        "0x1 Iconi") != "OK")
                        return false
                    this.DownSheetAsync(loc)
                } else {
                    if (MsgBox("ERP 작업보고를 시작합니다.`n`n장소 : " . loc . " (엑셀 수동입력)`n점검자 : " . memberStr, "진행합니다",
                        "0x1 Icon!") != "OK")
                        return false
                    this.OpenLocalExcel(loc)
                }
            } else {
                if (MsgBox("ERP 작업보고를 시작합니다.`n`n장소 : " . loc . "`n점검자 : " . memberStr, "진행합니다", "0x1 Iconi") != "OK")
                    return false
            }

            return this.Macro(memberStr, loc, userID, userPW, type, order, false)
        }
    }

    static GetMemberStr(members, format) {
        memberStr := ""
        if (members.Length > 0) {
            if (format == "summary") {
                memberStr := members[1] " 외 " (members.Length - 1) "명"
            } else {
                for index, name in members {
                    memberStr .= (index == 1 ? "" : ", ") . name
                }
            }
        }
        return memberStr
    }

    ; --------------------------------------------------------------------------
    ; Web / Excel Logic
    ; --------------------------------------------------------------------------
    static BatchDownSheetAsync(locations) {
        global XlsxGasURL, TARGET_SPREADSHEET_ID
        cmds := ""

        for index, ss in locations {
            ec_ss := URLEncode(ss)
            url := XlsxGasURL . "?fileId=" . TARGET_SPREADSHEET_ID . "&sheetName=" . ec_ss . "&filename=" . ec_ss .
                ".xlsx"
            tempFile := A_WorkingDir . "\temp_" . ss . ".json"
            localFile := A_WorkingDir . "\" . ss . ".xlsx"

            if FileExist(tempFile)
                FileDelete(tempFile)
            if FileExist(localFile)
                FileDelete(localFile)

            curlCmd := 'curl -sL --ssl-no-revoke -o "' . tempFile . '" "' . url . '"'
            if (index == 1)
                cmds := curlCmd
            else
                cmds .= " & " . curlCmd
        }

        if (cmds != "")
            Run(A_ComSpec ' /c ' cmds, , "Hide")
    }

    static DownSheetAsync(ss) {
        global XlsxGasURL, TARGET_SPREADSHEET_ID

        try {
            ec_ss := URLEncode(ss)
            url := XlsxGasURL . "?fileId=" . TARGET_SPREADSHEET_ID . "&sheetName=" . ec_ss . "&filename=" . ec_ss .
                ".xlsx"

            tempFile := A_WorkingDir . "\temp_" . ss . ".json"
            if FileExist(tempFile)
                FileDelete(tempFile)

            localFile := A_WorkingDir . "\" . ss . ".xlsx"
            if FileExist(localFile)
                FileDelete(localFile)

            Run 'curl -sL --ssl-no-revoke -o "' . tempFile . '" "' . url . '"', , "Hide"
            return true
        } catch {
            return false
        }
    }

    ; --------------------------------------------------------------------------
    ; File Processing (Decode & Save)
    ; --------------------------------------------------------------------------
    static ProcessDownload(ss) {
        tempFile := A_WorkingDir . "\temp_" . ss . ".json"
        targetFile := A_WorkingDir . "\" . ss . ".xlsx"
        tempTarget := A_WorkingDir . "\" . ss . "_writing.xlsx"

        if !FileExist(tempFile) {
            LogDebug("ProcessDownload [" ss "]: tempFile 없음 → 다운로드 미완료")
            return false
        }

        try {
            fileContent := ""
            try {
                fileContent := FileRead(tempFile, "UTF-8")
            } catch {
                LogDebug("ProcessDownload [" ss "]: tempFile 읽기 실패 (락 또는 쓰기 중)")
                return false
            }

            if (fileContent == "") {
                LogDebug("ProcessDownload [" ss "]: tempFile 비어있음")
                return false
            }

            data := JSON.parse(fileContent)

            if (data.Has("error")) {
                LogDebug("ProcessDownload [" ss "]: 서버 오류 → " data["error"])
                MsgBox("서버 오류: " . data["error"], "오류", "iconx")
                try FileDelete(tempFile)
                return false
            }

            if (!data.Has("base64")) {
                LogDebug("ProcessDownload [" ss "]: base64 키 없음 (JSON 구조 이상)")
                return false
            }

            binaryData := this.BufferFromBase64(data["base64"])
            LogDebug("ProcessDownload [" ss "]: 디코딩 완료, 크기=" binaryData.Size " bytes")

            ; 원자적 쓰기: 임시 파일에 먼저 쓰고 완료 후 rename
            f := FileOpen(tempTarget, "w")
            f.RawWrite(binaryData)
            f.Close()

            if FileExist(targetFile)
                FileDelete(targetFile)
            FileMove(tempTarget, targetFile)

            LogDebug("ProcessDownload [" ss "]: xlsx 저장 완료")
            try FileDelete(tempFile)
            return true

        } catch as e {
            LogDebug("ProcessDownload [" ss "]: 예외 → " e.Message)
            if FileExist(tempTarget)
                try FileDelete(tempTarget)
            return false
        }
    }

    ; --------------------------------------------------------------------------
    ; 다운로드 완료 폴링 대기 (비동기 curl이 tempFile을 완성할 때까지)
    ; --------------------------------------------------------------------------
    static WaitForDownload(ss, timeoutSec := 30) {
        tempFile := A_WorkingDir . "\temp_" . ss . ".json"
        maxTicks := (timeoutSec * 1000) // 200

        loop maxTicks {
            if FileExist(tempFile) {
                try {
                    if FileGetSize(tempFile) > 200
                        return true
                }
            }
            Sleep 200
        }

        LogDebug("WaitForDownload [" ss "]: " timeoutSec "초 타임아웃")
        return false
    }

    static BufferFromBase64(str) {
        if (str == "")
            return Buffer(0)

        ; CRYPT_STRING_BASE64 = 0x00000001
        size := 0
        DllCall("Crypt32\CryptStringToBinaryW", "Str", str, "UInt", 0, "UInt", 1, "Ptr", 0, "UInt*", &size, "Ptr", 0,
            "Ptr", 0)

        buf := Buffer(size)
        DllCall("Crypt32\CryptStringToBinaryW", "Str", str, "UInt", 0, "UInt", 1, "Ptr", buf, "UInt*", &size, "Ptr", 0,
            "Ptr", 0)

        return buf
    }

    static OpenLocalExcel(ss) {
        ; 엑셀 파일 열고 사용자 확인 대기
        try {
            Run(ss . ".xlsx", , "Max")
            if !WinWaitActive(ss, , 5) {
                ; 창 제목 매칭이 안 될 수도 있으니 관대하게 넘어감
            }

            MsgBox("측정값을 확인/수정 후 엑셀을 저장하고 종료해주세요.`n`n엑셀이 종료되면 자동으로 다음 단계(SAP 입력)가 진행됩니다.", "안내", "iconi")

            ; 엑셀 프로세스가 닫힐 때까지 대기 (Excel 파일명 윈도우)
            ; 정확한 핸들링을 위해 WinWaitClose 사용
            WinWaitClose(ss)
            Sleep 500
        } catch {
            MsgBox("엑셀 처리 중 오류 발생", "오류", "iconx")
        }
    }

    ; --------------------------------------------------------------------------
    ; SAP Automation Logic
    ; --------------------------------------------------------------------------
    static Macro(member, ss, uID, uPW, targetType, targetOrder, batchMode) {

        chk1 := true
        chk2 := true
        chk3 := true

        ; 이미 실행 중인 SAP가 존재하면 창이름 저장
        existingSapWindows := Map()
        for hwnd in WinGetList("ahk_class SAP_FRONTEND_SESSION")
            existingSapWindows[hwnd] := WinGetTitle("ahk_id " hwnd)

        ; SAP 실행 (이미 실행 중이면 활성화됨)
        try {
            Run("작업보고.sap")
        } catch {
            MsgBox("작업보고.sap 실행 파일을 찾을 수 없습니다.", "오류", "iconx")
            return false
        }

        loop 150 { ; SAP 진입 대기 (약 15초)
            Sleep 100

            ; 1. SAP GUI 보안 경고 처리
            if WinExist("SAP GUI 보안") and chk1 {
                Sleep 250
                ControlSend "{Space}", "Button1", "SAP GUI 보안" ; 허용
                Sleep 250
                ControlSend "{Enter}", "Button2", "SAP GUI 보안"
                chk1 := false
            }

            ; 2. 기존 로그인 창 (#32770) 처리
            login_hwnd := WinExist("작업완료보고 ahk_class #32770")
            if login_hwnd and chk2 {
                cUIA := UIA.ElementFromHandle(login_hwnd)
                cUIA.FindElement({ AutomationId: "1004" }).value := uID
                cUIA.FindElement({ AutomationId: "1005" }).value := uPW
                cUIA.FindElement({ AutomationId: "1" }).Invoke()
                Sleep 250
                chk1 := false
                chk2 := false
            }

            ; 3. 메인 세션 창 처리
            if WinExist("SAP ahk_class SAP_FRONTEND_SESSION", , "Easy") and chk3 {
                WinActivate
                Sleep 250
                Send "{Ctrl down}a{Ctrl up}" . uID
                Send "{Tab}"
                Send "{Raw}" . uPW
                Send "{Left}{Enter}"
                chk1 := false
                chk2 := false
                chk3 := false
            }

            ; 4. 오더번호 입력 창 확인
            targetHwnd := 0

            for hwnd in WinGetList("작업완료보고 ahk_class SAP_FRONTEND_SESSION") {
                currentTitle := WinGetTitle("ahk_id " hwnd)

                if (!existingSapWindows.Has(hwnd)
                    || existingSapWindows[hwnd] != currentTitle) {
                    if targetHwnd {
                        MsgBox("이번 실행의 작업완료보고 창을 하나로 특정할 수 없습니다.", "오류", "iconx")
                        return false
                    }
                    targetHwnd := hwnd
                }
            }

            if targetHwnd {
                sapWindow := "ahk_id " targetHwnd
                WinActivate(sapWindow)
                Sleep 750

                if (WinGetMinMax(sapWindow) != 1) {
                    WinMaximize(sapWindow)
                    Sleep 500
                }

                ; 오더번호 입력
                orderNum := targetOrder
                if (orderNum == "") {
                    MsgBox("해당 장소(" . ss . ")의 오더번호를 찾을 수 없습니다.`n설정을 확인해주세요.", "오류", "iconx")
                    return false
                }

                sapSecurityKey := "HKEY_CURRENT_USER\Software\SAP\SAPGUI Front\SAP Frontend Server\Security"

                try {
                    if (RegRead(sapSecurityKey, "WarnOnAttach", 1) != 0)
                        RegWrite(0, "REG_DWORD", sapSecurityKey, "WarnOnAttach")
                } catch as e {
                    MsgBox("SAP 스크립트 접근 알림 설정을 변경하지 못했습니다.`n" e.Message, "오류", "iconx")
                    return false
                }

                ; 스크립팅 객체에 연결
                sapRot := ComObject("SapROTWr.SapROTWrapper")

                ; 스크립팅 시작
                stage := "GetROTEntry"
                try {
                    sapAuto := sapRot.GetROTEntry("SAPGUI")

                    stage := "GetScriptingEngine"
                    sapApp := sapAuto.GetScriptingEngine()

                    ;다중 세션 환경에서 실제 열린 세션에 연결
                    sapSession := ""
                    Loop sapApp.Children.Count {
                        sapConnection := sapApp.Children.ElementAt(A_Index - 1)

                        Loop sapConnection.Children.Count {
                            candidate := sapConnection.Children.ElementAt(A_Index - 1)

                            try {
                                candidateHwnd := candidate.FindById("wnd[0]").Handle

                                ; SAP Handle은 32비트 Long이므로 하위 32비트끼리 비교
                                if ((candidateHwnd & 0xFFFFFFFF) = (targetHwnd & 0xFFFFFFFF)) {
                                    sapSession := candidate
                                    break
                                }
                            } catch {
                                ; 접근할 수 없는 세션은 건너뜀
                            }
                        }

                        if IsObject(sapSession)
                            break
                    }

                    if !IsObject(sapSession) {
                        MsgBox("작업완료보고 창에 해당하는 SAP 세션을 찾지 못했습니다.", "오류", "iconx")
                        return false
                    }

                    stage := "오더번호 입력"
                    sapSession.FindById("wnd[0]/usr/ctxtGS_0100-AUFNR").Text := String(orderNum)

                    stage := "Enter 입력"
                    sapSession.FindById("wnd[0]").SendVKey(0)
                } catch as e {
                    MsgBox("실패 단계: " stage "`n코드 행: " e.Line "`n오류: " e.Message,
                        "SAP 스크립팅 오류", "iconx")
                    return false
                }
                break ; 루프 탈출 -> 다음 단계
            }

            if (A_Index == 150) {
                MsgBox("시간초과: SAP 실행 실패", "오류", "iconx")
                return false
            }
        }


        ;작업보고 대기
        /*
        sleep 250
        CoordMode "Pixel", "Screen"
        GetCaretPos(&cx, &cy, &cw, &ch)
        nowColor := PixelGetColor(cx + 5, cy + 5)
        while nowColor != 0xDFEBF5 {
            WinActivate("작업완료보고")
            sleep 100
            send "{end}"
            if A_Index > 30 {
                MsgBox("타임아웃 - 작업보고 진입실패`n프로그램이 종료됩니다" getPos, "오류", "iconx")
                ExitApp
            }
            sleep 150
            GetCaretPos(&cx, &cy, &cw, &ch)
            nowColor := PixelGetColor(cx + 5, cy + 5)
            getPos := "`n좌표 " cx ", " cy " => 색상 : " PixelGetColor(cx + 5, cy + 5)
        }
        CoordMode "Pixel", "Client"
        sleep 500

        */

        ; 작업자 입력칸 대기
        workerId := "wnd[0]/usr/subSUB_CON:SAPMZPM2418:0110/tabsTS_0110/tabpTAB1/ssubSUB_CON01:SAPMZPM2418:0111/tblSAPMZPM2418TC_0101/txtGT_AFRUD-LTXA1[10,0]"
        reportReady := false

        Loop 60 {  ; 최대 약 15초
            try {
                if !sapSession.Busy {
                    workerField := sapSession.FindById(workerId)
                    reportReady := true
                    break
                }
            } catch {
                ; 화면 전환 중에는 아직 객체가 없을 수 있음
            }
            Sleep 250
        }

        if !reportReady {
            MsgBox("시간초과 - " ss " 작업보고 진입 확인에 실패했습니다.", "오류", "iconx")
            return false
        }

        workerField.Text := member

        /*
        ; 입력 시작
        Send "{Tab 16}"
        Sleep 250

        A_Clipboard := member
        Send "^v" ; 작업자 붙여넣기
        Sleep 500

        */

        ; 변전소인 경우 측정값 입력 진행
        if (targetType == "변전소") {
            /*
            ; 측정값 입력 (Shift+Tab으로 이동 후 입력)
            Send "{Shift down}{Tab 14}{Shift up}{Right 2}{Enter}"
            Sleep 500
            Send "{Tab 5}{Enter}" ; 업로드 버튼
            Sleep 500
            */

            ;측정값 입력 탭
            tab4Id := "wnd[0]/usr/subSUB_CON:SAPMZPM2418:0110/tabsTS_0110/tabpTAB4"
            ;엑셀 업로드 버튼
            uploadButtonId := tab4Id "/ssubSUB_CON01:SAPMZPM2418:0114/btn%#AUTOTEXT018"

            try {
                sapSession.FindById(tab4Id).Select()
                sapSession.FindById(uploadButtonId).SetFocus()
                Send "{Enter}" ; 업로드 버튼
            } catch as e {
                MsgBox("측정값 업로드 버튼 실행 실패:`n" e.Message, "오류", "iconx")
                return false
            }

            ; 파일 선택 창 대기
            if WinWait("열기 ahk_exe saplogon.exe", , 15) {
                Sleep 250
                Send "{Tab}{Shift down}{Tab}{Shift up}" ; 파일명 입력칸 포커스
                Sleep 250

                ; 파일 경로 입력
                localFile := A_WorkingDir . "\" . ss . ".xlsx"
                loop 20 {
                    ; 1. 파일이 이미 있고 최신이면 OK
                    if FileExist(localFile) {
                        if SubStr(FileGetTime(localFile, "M"), 1, 8) = FormatTime(, "yyyyMMdd") {
                            break
                        }
                    }

                    ; 2. 임시 파일 확인 및 변환 시도
                    if (this.ProcessDownload(ss)) {
                        break ; 변환 성공 (이제 Loop 다시 돌면 1번 조건 만족)
                    }

                    if (A_Index == 60) {
                        MsgBox("점검데이터 다운로드에 실패하였습니다`n처음부터 다시 시도하시기 바랍니다", "타임아웃", "iconx")
                        return ; 매크로 중단
                    }
                    Sleep 500
                }

                send localFile
                sleep 250
                send "{Enter}"
            }
            else {
                MsgBox("시간초과로 종료합니다 - 불러오기 실패", "타임아웃", "iconx")
                return false
            }

            ;입력확인
            /*
            sleep 250
            CoordMode "Pixel", "Screen"
            while !GetCaretPos(&cx, &cy, &cw, &ch) || PixelGetColor(cx + 5, cy + 5) != 0xFEF09E {
                WinActivate("작업완료보고")
                sleep 250

                if WinExist("SAP GUI 보안") {
                    WinActivate
                    sleep 250
                    controlsend("{Space}", "button1", "SAP GUI 보안")
                    sleep 250
                    controlsend("{Enter}", "button2", "SAP GUI 보안")
                }

                if WinExist("Microsoft Office Excel ahk_exe EXCEL.EXE")	;엑셀경고
                {
                    WinActivate
                    sleep 250
                    send "y"
                    sleep 250
                }

                send "{end}"
                if A_Index > 40 {
                    MsgBox("시간초과로 종료합니다 - 측정값 입력 실패", "타임아웃", "iconx")
                    return false
                }

            }
            CoordMode "Pixel", "Client"
            */

            measureId := "wnd[0]/usr/subSUB_CON:SAPMZPM2418:0110/tabsTS_0110/tabpTAB4/ssubSUB_CON01:SAPMZPM2418:0114/tblSAPMZPM2418TC_0104/txtGT_IMPTT-RDCNT[5,0]"
            valueLoaded := false
            deadline := A_TickCount + 15000  ; 15초제한

            while (A_TickCount < deadline) {
                try {

                    if !sapSession.Busy {
                        measuredValue := Trim(sapSession.FindById(measureId).Text)
                        if (measuredValue != "") {
                            valueLoaded := true
                            break
                        }
                    }

                    if (A_TickCount >= deadline)
                        break

                } catch {
                    ; 화면 전환 중 객체가 아직 없을 수 있음
                }

                Sleep 250
            }

            if !valueLoaded {
                MsgBox("시간초과 - 측정값 입력을 확인하지 못했습니다.", "오류", "iconx")
                return false
            }
        }

        if !batchMode
            MsgBox("입력이 완료되었습니다.`nERP 화면을 확인 후 저장하시기 바랍니다.", "완료", "iconi")
        else {
            send "^s"
            WinWait "SAP Easy Access  -  사용자 메뉴"
        }
        return true
    }
    ; --------------------------------------------------------------------------
    ; Firestore Realtime FormList
    ;
    ; 데이터 흐름:
    ; 1) ui/app.js가 Firestore publicCache/formList를 onSnapshot으로 구독합니다.
    ; 2) 문서가 변경되면 WebView2 메시지(updateERPFormList)로 Main.ahk에 전달됩니다.
    ; 3) Main.ahk가 이 클래스의 ApplyFormList()를 호출합니다.
    ; 4) 여기서는 오늘 수정된 시트만 골라 ValidLocations와 화면 상태를 갱신합니다.
    ;
    ; 이전 StartPolling/RequestStatus 방식과 달리 이 구간에서는 GAS를 주기적으로
    ; 호출하지 않습니다. 네트워크 구독은 ui/app.js가 담당하고, AHK는 전달받은
    ; 최신 목록을 보관하여 필터링과 날짜 변경 시의 로컬 재계산만 담당합니다.
    ; --------------------------------------------------------------------------

    ; Firestore에서 마지막으로 전달받은 전체 FormList입니다.
    ; 자정이 지나면 서버에 다시 요청하지 않고 이 목록을 사용해 상태를 재계산합니다.
    static FormListItems := []

    ; 실시간 상태 초기화와 자정 타이머가 중복 등록되는 것을 막는 플래그입니다.
    static RealtimeStatusStarted := false

    ; SetTimer에 같은 콜백 객체를 전달해 기존 예약을 해제하고 다시 등록할 수 있도록
    ; 바인딩한 콜백을 정적 변수에 보관합니다.
    static MidnightRefreshCallback := ""

    ; WebView와 메시지 수신 준비가 끝난 뒤 실시간 상태 처리를 한 번만 시작합니다.
    ; Firestore 연결을 직접 시작하는 함수는 아니며, ui/app.js의 구독 결과가 오기
    ; 전에도 현재 보관된 목록(초기에는 빈 배열)으로 화면 상태를 안전하게 초기화합니다.
    static StartRealtimeStatus() {
        if (this.RealtimeStatusStarted)
            return

        this.RealtimeStatusStarted := true

        ; 자정 경계에서 날짜 필터를 다시 적용하도록 1회성 타이머를 예약합니다.
        this.ScheduleMidnightRefresh()

        ; 최초 표시 시 남아 있을 수 있는 상태를 현재 날짜 기준으로 정리합니다.
        this.RefreshStatusForToday()
    }

    ; Main.ahk가 WebView2의 updateERPFormList 메시지를 받은 뒤 호출하는 진입점입니다.
    ; 전달된 스냅샷 전체를 교체 저장하고, 기존 목록 필터링 규칙을 즉시 다시 적용합니다.
    ; 이 함수에서는 curl, GAS GET 요청 또는 반복 타이머를 사용하지 않습니다.
    static ApplyFormList(items) {
        ; 예상하지 못한 메시지 때문에 기존 정상 목록이 훼손되지 않도록 배열만 받습니다.
        if !(items is Array) {
            LogDebug("[ERP FormList] Firestore 목록 형식이 올바르지 않음")
            return false
        }

        ; Firestore 스냅샷은 목록 전체의 최신 상태이므로 이전 캐시를 통째로 교체합니다.
        this.FormListItems := items

        ; 초기화보다 실시간 메시지가 먼저 도착한 경우에도 자정 타이머를 빠짐없이 설정합니다.
        if (!this.RealtimeStatusStarted)
            this.StartRealtimeStatus()

        ; 새 목록을 받은 즉시 오늘 날짜 기준의 유효 사업장과 화면 표시를 갱신합니다.
        this.RefreshStatusForToday()
        LogDebug("[ERP FormList] Firestore 실시간 목록 반영: " items.Length "건")
        return true
    }

    ; 마지막으로 수신한 FormList를 오늘 날짜 기준으로 로컬에서 필터링합니다.
    ; 기존 RequestStatus가 하던 핵심 규칙인 "오늘 수정된 sheetName만 유효"를 유지하되,
    ; 서버 요청 없이 ValidLocations와 WebView 화면 상태를 한 번에 다시 만듭니다.
    static RefreshStatusForToday() {
        global wv

        ; 이전 날짜나 삭제된 항목이 남지 않도록 두 Map을 매번 빈 상태에서 재구성합니다.
        statusMap := Map()
        this.ValidLocations := Map()
        todayStr := FormatTime(, "yyyy-MM-dd")

        for item in this.FormListItems {
            ; 필수 필드가 없는 비정상 항목은 전체 갱신을 중단하지 않고 건너뜁니다.
            if !(item is Map) || !item.Has("lastModifiedDate") || !item.Has("sheetName")
                continue

            ; Firebase가 미리 계산한 서울 날짜를 우선 사용하고, 구형 데이터와의 호환을
            ; 위해 없을 때만 원본 lastModifiedDate에서 오늘 날짜 문자열을 확인합니다.
            itemDate := item.Has("seoulDate") ? item["seoulDate"] : item["lastModifiedDate"]
            if (itemDate == todayStr || InStr(itemDate, todayStr)) {
                name := Trim(item["sheetName"])
                if (name == "")
                    continue

                ; statusMap은 UI 표시용, ValidLocations는 이후 XLSX 처리 가능 여부 확인용입니다.
                statusMap[name] := true
                this.ValidLocations[name] := true
            }
        }

        ; 계산된 결과만 WebView에 전달합니다. Firestore 또는 GAS로 나가는 요청은 없습니다.
        payload := Map("type", "updateERPStatus", "status", statusMap)
        try wv.PostWebMessageAsJson(JSON.stringify(payload))
    }

    ; 다음 로컬 자정 직후에 HandleMidnightRefresh()가 한 번 실행되도록 예약합니다.
    ; 매분/5분 반복 타이머가 아니라, 실행될 때마다 다음 자정을 다시 계산하는 1회성
    ; 타이머이므로 날짜가 바뀌지 않는 동안에는 아무 처리도 하지 않습니다.
    static ScheduleMidnightRefresh() {
        if (!this.MidnightRefreshCallback)
            this.MidnightRefreshCallback := ObjBindMethod(this, "HandleMidnightRefresh")

        ; 같은 콜백에 남아 있을 수 있는 기존 예약을 먼저 해제하여 중복 실행을 방지합니다.
        SetTimer this.MidnightRefreshCallback, 0
        todayStart := FormatTime(, "yyyyMMdd") . "000000"
        nextMidnight := DateAdd(todayStart, 1, "Days")
        secondsUntilMidnight := DateDiff(nextMidnight, A_Now, "Seconds")

        ; 시스템 시각 경계 오차로 전날로 계산되는 일을 피하려고 자정 1초 뒤 실행합니다.
        delayMs := Max(1000, secondsUntilMidnight * 1000 + 1000)
        SetTimer this.MidnightRefreshCallback, -delayMs
    }

    ; 날짜가 바뀌면 보관 중인 FormList만으로 오늘 상태를 다시 계산한 뒤,
    ; 다음 날 자정 실행을 다시 예약합니다. 이 과정에서도 네트워크 요청은 없습니다.
    static HandleMidnightRefresh() {
        this.RefreshStatusForToday()
        this.ScheduleMidnightRefresh()
        LogDebug("[ERP FormList] 날짜 변경으로 로컬 상태 재계산")
    }
}

GetCaretPos(&X, &Y, &W, &H) {
    /*
    	This implementation prefers CaretGetPos > Acc > UIA. This is mostly due to speed differences
    	between the methods and statistically it seems more likely that the UIA method is required the
    	least (Chromium apps support Acc as well).
    */
    ; Default caret
    savedCaret := A_CoordModeCaret
    CoordMode "Caret", "Screen"
    CaretGetPos(&X, &Y)
    CoordMode "Caret", savedCaret
    if IsInteger(X) and ((X | Y) != 0) {
        W := 4, H := 20
        return true
    }

    ; Acc caret
    static _ := DllCall("LoadLibrary", "Str", "oleacc", "Ptr")
    try {
        idObject := 0xFFFFFFF8 ; OBJID_CARET
        if DllCall("oleacc\AccessibleObjectFromWindow", "ptr", WinExist("A"), "uint", idObject &= 0xFFFFFFFF
        , "ptr", -16 + NumPut("int64", idObject == 0xFFFFFFF0 ? 0x46000000000000C0 : 0x719B3800AA000C81, NumPut("int64",
            idObject == 0xFFFFFFF0 ? 0x0000000000020400 : 0x11CF3C3D618736E0, IID := Buffer(16)))
        , "ptr*", oAcc := ComValue(9, 0)) = 0 {
            x := Buffer(4), y := Buffer(4), w := Buffer(4), h := Buffer(4)
            oAcc.accLocation(ComValue(0x4003, x.ptr, 1), ComValue(0x4003, y.ptr, 1), ComValue(0x4003, w.ptr, 1),
            ComValue(0x4003, h.ptr, 1), 0)
            X := NumGet(x, 0, "int"), Y := NumGet(y, 0, "int"), W := NumGet(w, 0, "int"), H := NumGet(h, 0, "int")
            if (X | Y) != 0
                return true
        }
    }

    ; UIA caret
    static IUIA := ComObject("{e22ad333-b25f-460c-83d0-0581107395c9}", "{34723aff-0c9d-49d0-9896-7ab52df8cd8a}")
    try {
        ComCall(8, IUIA, "ptr*", &FocusedEl := 0) ; GetFocusedElement
        /*
        	The current implementation uses only TextPattern GetSelections and not TextPattern2 GetCaretRange.
        	This is because TextPattern2 is less often supported, or sometimes reports being implemented
        	but in reality is not. The only downside to using GetSelections is that when text
        	is selected then caret position is ambiguous. Nevertheless, in those cases it most
        	likely doesn't matter much whether the caret is in the beginning or end of the selection.

        	If GetCaretRange is needed then the following code implements that:
        	ComCall(16, FocusedEl, "int", 10024, "ptr*", &patternObject:=0), ObjRelease(FocusedEl) ; GetCurrentPattern. TextPattern2 = 10024
        	if patternObject {
        		ComCall(10, patternObject, "int*", &IsActive:=1, "ptr*", &caretRange:=0), ObjRelease(patternObject) ; GetCaretRange
        		ComCall(10, caretRange, "ptr*", &boundingRects:=0), ObjRelease(caretRange) ; GetBoundingRectangles
        		if (Rect := ComValue(0x2005, boundingRects)).MaxIndex() = 3 { ; VT_ARRAY | VT_R8
        			X:=Round(Rect[0]), Y:=Round(Rect[1]), W:=Round(Rect[2]), H:=Round(Rect[3])
        			return
        		}
        	}
        */
        ComCall(16, FocusedEl, "int", 10014, "ptr*", &patternObject := 0), ObjRelease(FocusedEl) ; GetCurrentPattern. TextPattern = 10014
        if patternObject {
            ComCall(5, patternObject, "ptr*", &selectionRanges := 0), ObjRelease(patternObject) ; GetSelections
            ComCall(4, selectionRanges, "int", 0, "ptr*", &selectionRange := 0) ; GetElement
            ComCall(10, selectionRange, "ptr*", &boundingRects := 0), ObjRelease(selectionRange), ObjRelease(
                selectionRanges) ; GetBoundingRectangles
            if (Rect := ComValue(0x2005, boundingRects)).MaxIndex() = 3 { ; VT_ARRAY | VT_R8
                X := Round(Rect[0]), Y := Round(Rect[1]), W := Round(Rect[2]), H := Round(Rect[3])
                return true
            }
        }
    }

    return false
}
