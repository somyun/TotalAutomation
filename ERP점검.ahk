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

        location := msg.Has("location") ? msg["location"] : ""
        members := msg.Has("members") ? msg["members"] : []
        format := msg.Has("format") ? msg["format"] : "summary"
        targetType := msg.Has("targetType") ? msg["targetType"] : ""
        targetOrder := msg.Has("targetOrder") ? msg["targetOrder"] : ""
        userID := msg.Has("ID") ? msg["ID"] : ""
        userPW := msg.Has("sapPW") ? msg["sapPW"] : ""

        if (userPW == "") {
            MsgBox("SAP PW 지정되지 않아 실행할 수 없습니다", "오류", "iconx")
            return false
        }

        if (location == "") {
            MsgBox("예외 발생 : 장소 미지정", "오류", "iconx")
            return false
        }

        ; 1. 멤버 문자열 조합
        memberStr := ""
        if (members.Length > 0) {
            if (format == "summary") {
                memberStr := members[1] " 외 " (members.Length - 1) "명"
            } else {
                for index, name in members {
                    memberStr .= (index == 1 ? "" : ", ") . name
                }
            }
        } else {
            MsgBox("예외 발생 : 작업자 미지정", "오류", "iconx")
            return false
        }

        ; 2. 점검 시작
        isSubstation := (targetType == "변전소")
        isWebMode := false

        if !batchMode {
            if (isSubstation) {
                if (this.ValidLocations.Has(location)) {
                    ; [GREEN] WEB 연동 가능
                    isWebMode := true
                    if (MsgBox("ERP 작업보고를 시작합니다.`n`n장소 : " . location . " (WEB앱 연동)`n점검자 : " . memberStr,
                        "진행합니다", "0x1 Iconi") != "OK") {
                        return false
                    }
                    ; 비동기 다운로드 시작 (SAP 실행되는 동안 백그라운드 다운로드)
                    this.DownSheetAsync(location)
                } else {
                    ; [RED] WEB 연동 불가 (또는 아직 작성 안됨) -> 수동 모드
                    if (MsgBox("ERP 작업보고를 시작합니다.`n`n장소 : " . location . " (엑셀 수동입력)`n점검자 : " . memberStr,
                        "진행합니다", "0x1 Icon!") != "OK") {
                        return false
                    }
                    ; 로컬 엑셀 열기 (사용자가 확인/수정 후 종료하면 계속 진행)
                    this.OpenLocalExcel(location)
                }
            } else {
                ; 변전소 외 (전기실 등)
                if (MsgBox("ERP 작업보고를 시작합니다.`n`n장소 : " . location . "`n점검자 : " . memberStr,
                    "진행합니다", "0x1 Iconi") != "OK") {
                    return false
                }
            }
        }
        ;일괄모드 시
        else {
            if isSubstation {
                if this.ValidLocations.Has(location) {
                    isWebMode := true
                    this.DownSheetAsync(location)
                }
                else {
                    MsgBox "웹앱에 " location "의 데이터가 저장되어있지 않습니다.`n ERP 작업보고 일괄모드를 중단합니다.", "진행불가", "0x1 Icon!"
                    return false
                }
            }
        }

        ; 3. SAP 자동 입력 실행
        ; isWebMode 플래그는 여기서 딱히 필요 없지만(파일 날짜로 체크하므로), 로직 흐름상 명확히 함
        return this.Macro(memberStr, location, userID, userPW, targetType, targetOrder, batchMode)
    }

    ; --------------------------------------------------------------------------
    ; Web / Excel Logic
    ; --------------------------------------------------------------------------
    static DownSheetAsync(ss) {
        global WebAppURL, TARGET_SPREADSHEET_ID

        try {
            ec_ss := URLEncode(ss)
            url := WebAppURL . "?fileId=" . TARGET_SPREADSHEET_ID . "&sheetName=" . ec_ss . "&filename=" . ec_ss .
                ".xlsx"

            ; 임시 JSON 파일 경로 (고정된 이름 사용 또는 ID 기반)
            ; 여기서는 간편함을 위해 working dir에 temp_ + location + .json 저장
            tempFile := A_WorkingDir . "\temp_" . ss . ".json"

            ; 기존 임시 파일 및 타겟 엑셀 파일 삭제
            if FileExist(tempFile)
                FileDelete(tempFile)

            localFile := A_WorkingDir . "\" . ss . ".xlsx"
            if FileExist(localFile)
                FileDelete(localFile)

            ; 비동기 실행 (Run) - JSON으로 다운로드됨
            ; -s: Silent, -L: Follow redirects, -o: Output file
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

        if !FileExist(tempFile)
            return false

        try {
            ; 1. JSON 읽기
            ; 파일이 쓰기 중일 수 있으므로 읽기 시도
            fileContent := ""
            try {
                fileContent := FileRead(tempFile, "UTF-8")
            } catch {
                return false ; 아직 다운로드 중이거나 락 걸림
            }

            if (fileContent == "")
                return false

            ; 2. 파싱 및 디코딩
            data := JSON.parse(fileContent)

            if (data.Has("error")) {
                MsgBox("서버 오류: " . data["error"], "오류", "iconx")
                try FileDelete(tempFile) ; 에러 파일 삭제
                return false ; 영구 실패지만 루프에서 계속 재시도하지 않도록 처리는 상위에서
            }

            if (!data.Has("base64")) {
                ; JSON 형식이 아님 (혹시라도 그냥 엑셀이 받아진 경우?)
                return false
            }

            base64Str := data["base64"]
            binaryData := this.BufferFromBase64(base64Str)

            ; 3. 엑셀 파일로 저장
            f := FileOpen(targetFile, "w")
            f.RawWrite(binaryData)
            f.Close()

            ; 4. 성공 시 임시 파일 삭제
            try FileDelete(tempFile)
            return true

        } catch as e {
            ; 파싱 에러 등은 아직 다운로드가 덜 되어서 그럴 수 있음
            return false
        }
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

            ; 4. 오더번호 입력 창 진입 확인
            if WinExist("작업완료보고 ahk_class SAP_FRONTEND_SESSION") {
                WinActivate
                Sleep 750
                if (WinGetMinMax("작업완료보고 ahk_class SAP_FRONTEND_SESSION") != 1) {
                    WinMaximize
                    Sleep 500
                }

                ; 오더번호 입력
                orderNum := targetOrder
                if (orderNum == "") {
                    MsgBox("해당 장소(" . ss . ")의 오더번호를 찾을 수 없습니다.`n설정을 확인해주세요.", "오류", "iconx")
                    return false
                }

                Send orderNum . "{Enter}"
                break ; 루프 탈출 -> 다음 단계
            }

            if (A_Index == 150) {
                MsgBox("시간초과: SAP 실행 실패", "오류", "iconx")
                return false
            }
        }

        {	;작업보고 대기
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
        }

        ; 입력 시작
        Send "{Tab 16}"
        Sleep 250

        A_Clipboard := member
        Send "^v" ; 작업자 붙여넣기
        Sleep 500

        ; 변전소인 경우 측정값 입력 진행
        if (targetType == "변전소") {
            ; 측정값 입력 (Shift+Tab으로 이동 후 입력)
            Send "{Shift down}{Tab 14}{Shift up}{Right 2}{Enter}"
            Sleep 500
            Send "{Tab 5}{Enter}" ; 업로드 버튼
            Sleep 500

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

                    if (A_Index == 20) {
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
    ; Polling Logic
    ; --------------------------------------------------------------------------

    static IsPolling := false

    static StartPolling() {
        ; 5분마다 상태 갱신 요청
        SetTimer () => ERP점검.RequestStatus(), 300000
        ; 시작 시 즉시 1회 실행
        ERP점검.RequestStatus()
    }

    static RequestStatus() {

        ; 폴링 중복 방지
        if (this.IsPolling) {
            return
        }

        this.IsPolling := true

        global WebAppURL
        url := WebAppURL . "?action=getFormList"
        tempFile := A_ScriptDir . "\erp_status_temp.json"

        ; 기존 파일 정리
        if FileExist(tempFile) {
            try FileDelete tempFile
        }

        ; 비동기(Non-blocking) 실행: 외부 프로세스에 위임
        ; -s: Silent, -o: Output file, -L: Follow redirects (GAS 필수)
        ; curl이 없으면 실패하겠지만 Windows 10/11은 기본 내장됨
        cmd := 'curl.exe -skL "' . url . '" -o "' . tempFile . '"'

        try {
            Run cmd, , "Hide"
        } catch {
            ; curl 실행 실패 시 (경로 문제 등) silent 하게 넘어갑니다.
            this.IsPolling := false
            return
        }

        ; 타이머 콜백 초기화 (최초 1회)
        if !HasProp(ERP점검, "TimerCallback") || !ERP점검.TimerCallback
            ERP점검.TimerCallback := ObjBindMethod(ERP점검, "CheckResponse")

        ; 결과 확인 타이머 시작 (0.2초 간격)
        ERP점검.CheckCount := 0
        SetTimer ERP점검.TimerCallback, 200
    }

    static CheckCount := 0
    static TimerCallback := ""

    static CheckResponse() {
        tempFile := A_ScriptDir . "\erp_status_temp.json"

        ; CheckCount가 없으면 초기화 (만약을 대비)
        if !HasProp(ERP점검, "CheckCount")
            ERP점검.CheckCount := 0

        ERP점검.CheckCount += 1

        ; 타임아웃 처리 (약 10초 = 50회)
        if (ERP점검.CheckCount > 50) {
            if (ERP점검.TimerCallback)
                SetTimer ERP점검.TimerCallback, 0 ; 타이머 중지

            if FileExist(tempFile)
                try FileDelete tempFile

            this.IsPolling := false
            return
        }

        if !FileExist(tempFile)
            return

        ; 파일 읽기 시도
        try {
            fileContent := FileRead(tempFile, "UTF-8")
            if (fileContent == "")
                return ; 아직 다 안 써진 경우

            ; JSON 파싱 시도
            data := JSON.parse(fileContent)

            ; --- 성공 시 처리 ---
            if (ERP점검.TimerCallback)
                SetTimer ERP점검.TimerCallback, 0 ; 타이머 중지

            try FileDelete tempFile

            ; 폴링 상태 해제
            this.IsPolling := false

            global wv
            statusMap := Map()
            todayStr := FormatTime(, "yyyy-MM-dd")

            if (data is Array) {
                ERP점검.ValidLocations := Map() ; 캐시 초기화
                for item in data {
                    if (!item.Has("lastModifiedDate") || !item.Has("sheetName"))
                        continue

                    lmDateRaw := item["lastModifiedDate"]
                    if InStr(lmDateRaw, todayStr) {
                        name := Trim(item["sheetName"]) ; 공백 제거
                        statusMap[name] := true
                        ERP점검.ValidLocations[name] := true ; 유효 목록 업데이트
                    }
                }
            }

            ; WebView로 전송
            payload := Map("type", "updateERPStatus", "status", statusMap)
            jsonStr := JSON.stringify(payload)
            wv.PostWebMessageAsJson(jsonStr)

        } catch as e {
            ; JSON 파싱 에러 등은 무시하고 다음 틱 재시도
        }
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
