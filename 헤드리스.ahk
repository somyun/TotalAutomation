#Requires AutoHotkey v2.0
#Include Lib\JSON.ahk

class HeadlessAutomation {

    static CookieStorage := Map() ; 도메인별 쿠키 저장소 (btcep, niw, ep)
    static CookieFile := A_ScriptDir "\.saved_cookies.json"

    static SetCookieStorage(data) {
        if (Type(data) == "Map") {
            this.CookieStorage := data
        } else {
            ; Object 타입으로 파싱되었을 경우 Map으로 안전하게 복사
            for k, v in data.OwnProps()
                this.CookieStorage[k] := v
        }

        ; 저장된 쿠키 파일 백업 (리로드 대응)
        this.SaveCookieStorage()

        ; 디버그 로그 (임시)
        try {
            logMsg := "HeadlessAutomation.SetCookieStorage 호출됨. Type=" Type(data)
            if (Type(data) == "Map")
                logMsg .= ", Count=" data.Count
            logMsg .= " -> 저장된 CookieStorage Count=" this.CookieStorage.Count

            FileAppend(FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") " [Headless] " logMsg "`n", A_ScriptDir "\debug_main.txt",
            "UTF-8")
        }
    }

    ; ==============================================================================
    ; 쿠키 파일 저장/불러오기 (리로드 시 유지)
    ; ==============================================================================
    static SaveCookieStorage() {
        try {
            if (this.CookieStorage.Count > 0) {
                FileOpen(this.CookieFile, "w", "UTF-8").Write(JSON.stringify(this.CookieStorage))
            }
        }
    }

    static LoadCookieStorage() {
        try {
            if FileExist(this.CookieFile) {
                content := FileRead(this.CookieFile, "UTF-8")
                if (content != "") {
                    data := JSON.parse(content)
                    if (Type(data) == "Map") {
                        this.CookieStorage := data
                    } else {
                        for k, v in data.OwnProps()
                            this.CookieStorage[k] := v
                    }
                }
            }
        }
    }

    ; ==============================================================================
    ; CDP(Chrome DevTools Protocol)용 쿠키 파라미터 생성
    ; ==============================================================================
    static GetCookieParamsForCDP() {
        cookieParams := []

        ; 저장소의 모든 쿠키를 CDP 형식으로 변환
        for domainKey, cookieStr in this.CookieStorage {
            loop parse, cookieStr, ";" {
                if (A_LoopField == "")
                    continue

                parts := StrSplit(A_LoopField, "=", , 2)
                if (parts.Length == 2) {
                    cName := Trim(parts[1])
                    cValue := Trim(parts[2])

                    ; 도메인 매핑
                    cDomain := "ep.humetro.busan.kr"
                    if (domainKey == "btcep")
                        cDomain := "btcep.humetro.busan.kr"
                    else if (domainKey == "niw")
                        cDomain := "niw.humetro.busan.kr"

                    cookieParams.Push(Map(
                        "name", cName,
                        "value", cValue,
                        "domain", cDomain,
                        "path", "/",
                        "url", "http://" cDomain "/"
                    ))
                }
            }
        }
        return cookieParams
    }

    ; ==========================================================================
    ; 생성자: HTTP 클라이언트 모드
    ; ==========================================================================
    __New(headless := true, logger := "") {
        this.logFunc := logger
    }

    ; Connect 메서드 (호환성 유지)
    static Connect(logger := "") {
        return HeadlessAutomation(true, logger)
    }

    _Log(msg) {
        if (this.logFunc)
            this.logFunc.Call(msg)
    }

    ; ==============================================================================
    ; 금일 일지 번호 조회 (HTTP)
    ; ==============================================================================
    GetTodayWorkLogNumber(sabun, deptCode, targetDate := "") {
        if (targetDate == "")
            targetDate := FormatTime(DateAdd(A_Now, -9, "Hours"), "yyyyMMdd")

        this._Log("일지 번호 조회 시작... (" targetDate ", " deptCode ")")

        ; 1. 날짜 포맷 변환 (YYYY-MM-DD -> YYYYMMDD)
        targetDateSimple := StrReplace(targetDate, "-", "")

        ; 2. Payload 구성
        ; WorkLogData는 x-www-form-urlencoded 방식을 사용함
        ; 2. Payload 구성 (Form Data String)
        ; WorkLogData는 x-www-form-urlencoded 방식을 사용함
        payloadStr := "start=0&limit=25&forumId=4&I_ARWRK=5010"
            . "&I_ARBPL=" deptCode
            . "&I_WERKS=5010"
            . "&I_GIJUNDF=" targetDateSimple
            . "&I_GIJUNDT=" targetDateSimple
            . "&I_SANGTAE="
            . "&I_ARBPL01=" deptCode

        loop 15 {
            payloadStr .= "&I_ARBPL" . Format("{:02}", A_Index + 1) . "="
        }

        url :=
            "http://ep.humetro.busan.kr/irj/servlet/prt/portal/prtroot/kr.busan.humetro.cbo.erp.work_log.WorkLogData"

        ; 3. 요청 전송 (Form Data 명시)
        responseText := this._SendRequest(payloadStr, url, "application/x-www-form-urlencoded; charset=UTF-8")

        if (responseText == "") {
            this._Log("일지 목록 조회 실패 (응답 없음)")
            return ""
        }

        try {
            resObj := JSON.parse(responseText)
            if (resObj.Has("ZPM_RFC_REPORT_ZTPM1000_LIST")) {
                list := resObj["ZPM_RFC_REPORT_ZTPM1000_LIST"]
                if (list.Length > 0) {
                    iljino := list[1]["ILJINO"]
                    this._Log("일지번호 발견: " iljino)
                    return iljino
                }
            }
            this._Log("일지 목록이 비어있거나 파싱 실패")
            return ""
        } catch as e {
            this._Log("일지번호 조회 중 오류: " e.Message)
            return ""
        }
    }

    ; ==============================================================================
    ; 종료 로직 (HTTP 모드에서는 특별한 정리 필요 없음)
    ; ==============================================================================
    Close(killBrowser := false) {
        ; No-op
    }

    ; ==============================================================================
    ; 공통 HTTP 요청 함수 (WinHttp API 사용)
    ; ==============================================================================

    _SendRequest(payload, url, contentType := "") {
        ; 1. URL 기반 도메인 판별 및 쿠키 선택
        domain := "ep" ; 기본값
        if InStr(url, "btcep.humetro.busan.kr")
            domain := "btcep"
        else if InStr(url, "niw.humetro.busan.kr")
            domain := "niw"

        cookie := ""
        if (HeadlessAutomation.CookieStorage.Has(domain))
            cookie := HeadlessAutomation.CookieStorage[domain]

        if (cookie == "") {
            this._Log("Error: " domain " 도메인 쿠키가 없습니다.")
            return ""
        }

        try {
            whr := ComObject("WinHttp.WinHttpRequest.5.1")

            ; 3초 타임아웃 설정 (Resolve, Connect, Send, Receive)
            whr.SetTimeouts(3000, 3000, 3000, 3000)

            ; 비동기 모드(true)로 열어야 WaitForResponse로 타임아웃 제어 가능
            whr.Open("POST", url, true)

            ; 2. Content-Type 결정 로직
            finalContentType := contentType
            finalPayload := payload

            if (finalContentType == "") {
                ; contentType 미지정 시, payload 타입에 따라 자동 결정
                if (IsObject(payload)) {
                    finalContentType := "application/json; charset=UTF-8" ; JSON (추정)
                    finalPayload := JSON.stringify(payload)
                } else {
                    finalContentType := "application/x-www-form-urlencoded; charset=UTF-8" ; Default
                }
            } else {
                ; contentType 지정 시, payload가 객체면 문자열 변환만 수행
                if (IsObject(payload)) {
                    finalPayload := JSON.stringify(payload)
                }
            }

            ; 헤더 설정
            whr.SetRequestHeader("Content-Type", finalContentType)
            whr.SetRequestHeader("Cookie", cookie)
            whr.SetRequestHeader("Accept-Language", "ko,en;q=0.9,en-US;q=0.8")
            whr.SetRequestHeader("X-Requested-With", "XMLHttpRequest") ; AJAX 필수 헤더 추가 - 요청 실패시 주석 해제 검토

            ; Payload 전송
            whr.Send(finalPayload)

            ; 3초 대기 (성공 시 -1 반환, 타임아웃 시 0 반환)
            if (whr.WaitForResponse(3)) {
                try {
                    ; ResponseBody 처리 (Binary -> UTF-8 String)
                    body := whr.ResponseBody

                    ; ComObjArray 처리
                    size := body.MaxIndex() + 1
                    buf := Buffer(size)

                    loop size
                        NumPut("UChar", body[A_Index - 1], buf, A_Index - 1)

                    responseText := StrGet(buf, "UTF-8")
                    return responseText

                } catch as e {
                    this._Log("Error: 응답 데이터 디코딩 실패 - " e.Message)
                    return ""
                }
            } else {
                this._Log("Error: 요청 시간 초과 (3초)")
                return ""
            }

        } catch as e {
            this._Log("Error: 통신 실패 (WinHttp) - " e.Message)
            return ""
        }
    }

    ; ==============================================================================
    ; 전체 근무자/근태 정보 조회
    ; ==============================================================================
    GetWorkerList(arbpl) {
        this._Log("전체 근무자 정보 조회 시작 (작업장: " arbpl ", HTTP)")

        ; 날짜 계산
        targetDate := FormatTime(A_Now, "yyyyMMdd")

        ; 08:30 이전 처리 (전일자 조회)
        if (FormatTime(A_Now, "HHmm") < "0830")
            targetDate := FormatTime(DateAdd(A_Now, -1, "days"), "yyyyMMdd")

        payload := Map(
            "AJAX_TYPE", [Map(
                "MODE", "OPN", "TYPE", "PSN",
                "I_ARBPL", arbpl, "I_WERKS", "5010", "I_GIJUND", targetDate
            )]
        )

        ; Referer: WorkLogPerson
        responseText := this._SendRequest(payload,
            "http://ep.humetro.busan.kr/irj/servlet/prt/portal/prtroot/kr.busan.humetro.cbo.erp.work_log.WorkLogDtlData",
            "application/x-www-form-urlencoded")

        if (responseText == "") {
            return []
        }

        workerRecords := []
        try {
            result := JSON.parse(responseText)
            if (result.Has("ZPM_RFC_REPORT_ZTPM1017_LIST")) {
                listData := result["ZPM_RFC_REPORT_ZTPM1017_LIST"]

                this._Log("데이터 파싱 성공: " listData.Length "명 조회됨")

                for item in listData {
                    ; 모든 인원 추가
                    empno := item.Has("PERNR") ? item["PERNR"] : ""
                    name := item.Has("SMNAM") ? item["SMNAM"] : ""
                    sagot := item.Has("SAGOT") ? item["SAGOT"] : ""
                    gubunt := item.Has("GUBUNT") ? item["GUBUNT"] : ""

                    if (empno != "") {
                        workerRecords.Push(Map(
                            "사번", empno,
                            "이름", name,
                            "휴가종류", sagot,
                            "근무조", gubunt
                        ))
                    }
                }
            } else {
                this._Log("응답 구조 상이함: ZPM_RFC_REPORT_ZTPM1017_LIST 없음")
            }
        } catch as e {
            this._Log("JSON 파싱 오류: " e.Message)
        }

        return workerRecords
    }

    ; ==============================================================================
    ; 오더 목록 조회
    ; ==============================================================================
    GetOrderList(arbpl) {
        this._Log("오더 목록 조회 요청 (" arbpl ", HTTP)...")

        ; 날짜 계산
        targetDate := FormatTime(A_Now, "yyyyMMdd")
        nextDay := DateAdd(A_Now, 1, "days")
        tomorrow := FormatTime(nextDay, "yyyyMMdd")

        ; 08:30 이전 처리 (전일자 조회)
        if (FormatTime(A_Now, "HHmm") < "0830") {
            targetDate := FormatTime(DateAdd(A_Now, -1, "days"), "yyyyMMdd")
            tomorrow := FormatTime(A_Now, "yyyyMMdd")
        }

        payload := Map(
            "AJAX_TYPE", [Map(
                "MODE", "OPN", "TYPE", "SPC", "RFC_NAME", "ZPM_RFC_REPORT_AFRU_LIST",
                "I_ARBPL", arbpl, "I_WERKS", "5010", "I_FRDAY", targetDate,
                "I_FRTIME", "090000", "I_TODAY", tomorrow, "I_TOTIME", "085900"
            )]
        )

        ; Referer: WorkLogSpec
        responseText := this._SendRequest(payload,
            "http://ep.humetro.busan.kr/irj/servlet/prt/portal/prtroot/kr.busan.humetro.cbo.erp.work_log.WorkLogDtlData",
            "application/x-www-form-urlencoded")

        if (responseText == "") {
            return []
        }

        aufnrList := []
        try {
            result := JSON.parse(responseText)
            if (result.Has("ZPM_RFC_REPORT_AFRU_LIST")) {
                listData := result["ZPM_RFC_REPORT_AFRU_LIST"]
                for item in listData {
                    if (item.Has("AUFNR"))
                        aufnrList.Push(SubStr(item["AUFNR"], 5))
                }
            }
        } catch as e {
            this._Log("JSON 파싱 오류: " e.Message)
        }

        return aufnrList
    }
}

class BackgroundProcessManager {
    static hReadPipe := 0  ; Stdout Read (Parent)
    static hWritePipe := 0 ; Stdout Write (Child - Closed in Parent)
    static hReadStdin := 0 ; Stdin Read (Child - Closed in Parent)
    static hWriteStdin := 0 ; Stdin Write (Parent)
    static hProcess := 0
    static LogCallback := ""

    ; 백그라운드 프로세스 실행 (Two-way Pipe)
    static Launch(cmdLine, callback) {
        this.Cleanup()
        this.LogCallback := callback

        sa := Buffer(24, 0) ; SECURITY_ATTRIBUTES (x64)
        NumPut("UInt", 24, sa, 0)
        NumPut("Ptr", 0, sa, 8)
        NumPut("Int", 1, sa, 16) ; bInheritHandle = TRUE

        ; 1. Stdout 파이프 생성 (Child Write -> Parent Read)
        hReadOut := 0, hWriteOut := 0
        if !DllCall("CreatePipe", "PtrP", &hReadOut, "PtrP", &hWriteOut, "Ptr", sa, "UInt", 0) {
            MsgBox("Stdout 파이프 생성 실패")
            return false
        }
        ; 읽기 핸들은 상속되지 않도록 설정 (선택사항이나 권장됨)
        DllCall("SetHandleInformation", "Ptr", hReadOut, "UInt", 1, "UInt", 0)

        this.hReadPipe := hReadOut
        this.hWritePipe := hWriteOut

        ; 2. Stdin 파이프 생성 (Parent Write -> Child Read)
        hReadIn := 0, hWriteIn := 0
        if !DllCall("CreatePipe", "PtrP", &hReadIn, "PtrP", &hWriteIn, "Ptr", sa, "UInt", 0) {
            MsgBox("Stdin 파이프 생성 실패")
            this.Cleanup()
            return false
        }
        ; 쓰기 핸들은 상속되지 않도록 설정
        DllCall("SetHandleInformation", "Ptr", hWriteIn, "UInt", 1, "UInt", 0)

        this.hReadStdin := hReadIn
        this.hWriteStdin := hWriteIn

        ; 3. 프로세스 생성
        si := Buffer(104, 0) ; STARTUPINFO (x64)
        NumPut("UInt", 104, si, 0)
        NumPut("UInt", 0x100, si, 60) ; dwFlags = STARTF_USESTDHANDLES
        NumPut("Ptr", this.hReadStdin, si, 80) ; hStdInput
        NumPut("Ptr", this.hWritePipe, si, 88) ; hStdOutput
        NumPut("Ptr", this.hWritePipe, si, 96) ; hStdError

        pi := Buffer(24, 0) ; PROCESS_INFORMATION

        if !DllCall("CreateProcess", "Ptr", 0, "Str", cmdLine, "Ptr", 0, "Ptr", 0, "Int", 1, "UInt", 0x08000000, "Ptr",
            0, "Ptr", 0, "Ptr", si, "Ptr", pi) {
            MsgBox("프로세스 생성 실패")
            this.Cleanup()
            return false
        }

        ; 부모 프로세스에서는 자식용 핸들 닫기 (중요: 그래야 자식 종료 시 EOF 감지 가능)
        DllCall("CloseHandle", "Ptr", this.hWritePipe)
        this.hWritePipe := 0
        DllCall("CloseHandle", "Ptr", this.hReadStdin)
        this.hReadStdin := 0

        this.hProcess := NumGet(pi, 0, "Ptr")
        hThread := NumGet(pi, 8, "Ptr")
        DllCall("CloseHandle", "Ptr", hThread)

        ; 모니터링 타이머 시작
        SetTimer ObjBindMethod(this, "CheckOutput"), 50
        return true
    }

    ; 표준 입력(Stdin)으로 데이터 전송
    static SendInput(text) {
        if (!this.hWriteStdin)
            return false

        ; UTF-8 변환 (개행 문자 추가 필수)
        if (SubStr(text, -1) != "`n")
            text .= "`n"

        bufLen := StrPut(text, "UTF-8")
        buf := Buffer(bufLen)
        StrPut(text, buf, "UTF-8")

        written := 0
        return DllCall("WriteFile", "Ptr", this.hWriteStdin, "Ptr", buf, "UInt", bufLen - 1, "UIntP", &written, "Ptr",
            0)
    }

    ; 출력 확인 (Non-blocking)
    static CheckOutput() {
        if (!this.hReadPipe) {
            SetTimer ObjBindMethod(this, "CheckOutput"), 0
            return
        }

        ; 1. 데이터 확인 (Peek)
        available := 0
        if !DllCall("PeekNamedPipe", "Ptr", this.hReadPipe, "Ptr", 0, "UInt", 0, "Ptr", 0, "UIntP", &available, "Ptr",
            0) {
            ; 파이프가 깨짐 (프로세스 종료 등)
            this.Cleanup()
            return
        }

        ; 2. 데이터 읽기
        if (available > 0) {
            bufsize := 8096 ; [개선] 버퍼 크기 증가 (1KB -> 8KB)
            buf := Buffer(bufsize, 0)
            read := 0

            if DllCall("ReadFile", "Ptr", this.hReadPipe, "Ptr", buf, "UInt", bufsize, "UIntP", &read, "Ptr", 0) {
                if (read > 0) {
                    ;text := StrGet(buf, read, "CP949")
                    text := StrGet(buf, read, "UTF-8")
                    lines := StrSplit(text, ["`r`n", "`n"])

                    if (this.LogCallback) {
                        for line in lines {
                            if (line != "")
                                this.LogCallback.Call(line)
                        }
                    }
                }
            }
        }

        ; 3. 프로세스 종료 확인 (적극적 확인)
        exitCode := 0
        if DllCall("GetExitCodeProcess", "Ptr", this.hProcess, "UIntP", &exitCode) {
            if (exitCode != 259) { ; STILL_ACTIVE
                this.Cleanup()
            }
        }
    }

    static Cleanup() {
        SetTimer ObjBindMethod(this, "CheckOutput"), 0

        if (this.hReadPipe) {
            DllCall("CloseHandle", "Ptr", this.hReadPipe)
            this.hReadPipe := 0
        }
        if (this.hWritePipe) {
            DllCall("CloseHandle", "Ptr", this.hWritePipe)
            this.hWritePipe := 0
        }
        if (this.hWriteStdin) {
            DllCall("CloseHandle", "Ptr", this.hWriteStdin)
            this.hWriteStdin := 0
        }
        if (this.hReadStdin) {
            DllCall("CloseHandle", "Ptr", this.hReadStdin)
            this.hReadStdin := 0
        }
        if (this.hProcess) {
            DllCall("CloseHandle", "Ptr", this.hProcess)
            this.hProcess := 0
        }
    }
}
