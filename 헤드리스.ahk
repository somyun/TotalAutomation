#Requires AutoHotkey v2.0
#Include Lib\Chrome.ahk
#Include Lib\JSON.ahk

class HeadlessAutomation {

    ChromeInst := ""
    PageInst := ""
    DebugPort := 9222
    epWorklogPath := "/irj/servlet/prt/portal/prtroot/kr.busan.humetro.cbo.erp.work_log.WorkLogDtlData"
    logFunc := ""

    ; ==========================================================================
    ; 생성자: 이미 실행된 브라우저(Port 9222)에 연결만 수행
    ; ==========================================================================
    __New(headless := true, logger := "") {
        this.logFunc := logger
        this._Log("브라우저 연결 시도 (Port: " this.DebugPort ")...")

        try {
            ; 중요: Chrome.ahk 초기화 시 'Chrome 실행 파일 경로'와 'headless=true'를 명시해야 함.
            ; 실행은 Python이 하지만, Chrome.ahk가 내부적으로 정보를 필요로 할 수 있음.
            chromePath := "C:\Program Files\Google\Chrome\Application\Chrome.exe"

            ; Connect Only Mode (프로필 경로 비워둠, Flags 비워둠)
            ; 4번째 인자: DebugPort, 6번째 인자: Headless 여부
            this.ChromeInst := Chrome("", "", chromePath, this.DebugPort, , headless)

            ; 페이지 인스턴스 확보
            try {
                this.PageInst := this.ChromeInst.GetPage()
            } catch {
                this.PageInst := this.ChromeInst.NewPage()
            }

            if !this.PageInst
                throw Error("페이지 인스턴스를 찾을 수 없습니다.")

            ; 봇 탐지 회피용 JS (안전장치)
            try {
                this.PageInst.Evaluate("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            }

        } catch as e {
            this._Log("브라우저 연결 실패: " e.Message)
            throw e
        }
    }

    ; Connect 메서드 (Main에서 호출됨)
    static Connect(logger := "") {
        return HeadlessAutomation(true, logger)
    }

    _Log(msg) {
        if (this.logFunc)
            this.logFunc.Call(msg)
    }

    ; ==============================================================================
    ; 종료 로직
    ; ==============================================================================
    Close(killBrowser := false) {
        if (this.ChromeInst) {
            this.ChromeInst := ""
        }
    }

    ; ==============================================================================
    ; 전체 근무자/근태 정보 조회
    ; ==============================================================================
    GetWorkerList(arbpl) {
        this._Log("전체 근무자 정보 조회 시작 (작업장: " arbpl ")")

        ; 날짜 계산
        targetDate := FormatTime(A_Now, "yyyyMMdd")

        ; 08:30 이전 처리 (전일자 조회)
        if (FormatTime(A_Now, "HHmm") < "0830")
            targetDate := FormatTime(DateAdd(A_Now, -1, "days"), "yyyyMMdd")

        dateStrToday := FormatTime(targetDate, "yyyyMMdd")

        jsCode := "
        (
            const payload = {
                AJAX_TYPE: [{
                    MODE: 'OPN', TYPE: 'PSN',
                    I_ARBPL: arbpl, I_WERKS: '5010', I_GIJUND: today
                }]
            };

            fetch(url, {
                method: 'POST', credentials: 'include',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8', 'X-Requested-With': 'XMLHttpRequest' },
                body: JSON.stringify(payload)
            })
            .then(r => r.text())
            .then(txt => { window._ahk_vac_res = txt; window._ahk_vac_done = true; })
            .catch(err => { window._ahk_vac_res = 'ERROR:' + err; window._ahk_vac_done = true; });
        )"

        RunJS := "window._ahk_vac_res = ''; window._ahk_vac_done = false; "
        RunJS .= "(function(arbpl, today, url) { " . jsCode . " })"
        RunJS .= "('" . arbpl . "', '" . dateStrToday . "', '" . this.epWorklogPath . "');"

        this.PageInst.Evaluate(RunJS)

        loop 40 {
            if (this.EvaluateSafe("window._ahk_vac_done") == 1)
                break
            Sleep 250
        }

        responseText := this.EvaluateSafe("window._ahk_vac_res")
        if (responseText == "" || InStr(responseText, "ERROR:")) {
            this._Log("근무자 조회 실패: " responseText)
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
        this._Log("오더 목록 조회 요청 (" arbpl ")...")

        ; 날짜 계산
        targetDate := FormatTime(A_Now, "yyyyMMdd")
        nextDay := DateAdd(A_Now, 1, "days")
        tomorrow := FormatTime(nextDay, "yyyyMMdd")

        ; 08:30 이전 처리 (전일자 조회)
        if (FormatTime(A_Now, "HHmm") < "0830") {
            targetDate := FormatTime(DateAdd(A_Now, -1, "days"), "yyyyMMdd")
            tomorrow := FormatTime(A_Now, "yyyyMMdd")
        }

        jsCode := "
        (
            const payload = {
                AJAX_TYPE: [{
                    MODE: 'OPN', TYPE: 'SPC', RFC_NAME: 'ZPM_RFC_REPORT_AFRU_LIST',
                    I_ARBPL: arbpl, I_WERKS: '5010', I_FRDAY: today,
                    I_FRTIME: '090000', I_TODAY: tomorrow, I_TOTIME: '085900'
                }]
            };

            fetch(url, {
                method: 'POST', credentials: 'include',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8', 'X-Requested-With': 'XMLHttpRequest' },
                body: JSON.stringify(payload)
            })
            .then(r => r.text())
            .then(txt => { window._ahk_ep_res = txt; window._ahk_ep_done = true; })
            .catch(err => { window._ahk_ep_res = 'ERROR:' + err; window._ahk_ep_done = true; });
        )"

        RunJS := "window._ahk_ep_res = ''; window._ahk_ep_done = false; "
        RunJS .= "(function(arbpl, today, tomorrow, url) { " . jsCode . " })"
        RunJS .= "('" . arbpl . "', '" . targetDate . "', '" . tomorrow . "', '" . this.epWorklogPath . "');"

        this.PageInst.Evaluate(RunJS)

        ; 결과 대기
        loop 40 {
            val := this.EvaluateSafe("window._ahk_ep_done")
            if (val == 1 || val == "true")
                break
            Sleep 250
        }

        responseText := this.EvaluateSafe("window._ahk_ep_res")

        if (responseText == "" || InStr(responseText, "ERROR:")) {
            this._Log("오더 조회 실패: " responseText)
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

    ; ==============================================================================
    ; [Helper] EvaluateSafe (Map 처리 포함)
    ; ==============================================================================
    EvaluateSafe(js) {
        try {
            res := this.PageInst.Evaluate(js)
            if (IsObject(res) && res.Has("value"))
                return res["value"]
            return String(res)
        } catch {
            return ""
        }
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
            bufsize := 1024
            buf := Buffer(bufsize, 0)
            read := 0

            if DllCall("ReadFile", "Ptr", this.hReadPipe, "Ptr", buf, "UInt", bufsize, "UIntP", &read, "Ptr", 0) {
                if (read > 0) {
                    text := StrGet(buf, read, "CP949")
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
