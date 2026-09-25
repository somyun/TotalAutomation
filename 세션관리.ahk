#Requires AutoHotkey v2.0
#Include "Lib\SessionAuth.ahk"

class SessionManager {
    static Session := 0
    static State := "empty" ; empty, acquiring, ready, error
    static CurrentUser := 0
    static ReadyCallbacks := []
    static ErrorCallbacks := []

    static AcquireAsync(user, onReady := 0, onError := 0) {
        if IsObject(onReady)
            this.ReadyCallbacks.Push(onReady)
        if IsObject(onError)
            this.ErrorCallbacks.Push(onError)

        if (this.State = "acquiring")
            return true
        if !IsObject(user) || !user.Has("id") || user["id"] = "" {
            this.Fail("로그인 사용자 정보가 없습니다.")
            return false
        }
        if !user.Has("webPW") || user["webPW"] = "" || !user.Has("pw2") || user["pw2"] = "" {
            this.Fail("통합 비밀번호 또는 2차 비밀번호가 등록되지 않았습니다.")
            return false
        }

        this.CurrentUser := user
        this.Session := 0
        this.State := "acquiring"
        this.NotifyState("acquiring")

        if A_IsCompiled {
            workerPath := A_ScriptDir "\세션BG.exe"
            commandLine := Format('"{1}"', workerPath)
        } else {
            workerPath := A_ScriptDir "\세션BG.ahk"
            commandLine := Format('"{1}" "{2}" --session-worker', A_AhkPath, workerPath)
        }

        if !FileExist(workerPath) {
            this.Fail("세션 백그라운드 실행 파일을 찾을 수 없습니다.")
            return false
        }

        if !BackgroundProcessManager.Launch(commandLine, ObjBindMethod(this, "HandleWorkerFrame")) {
            this.Fail("세션 백그라운드 프로세스를 실행하지 못했습니다.")
            return false
        }

        credentials := Map(
            "사번", user["id"],
            "1차비번", user["webPW"],
            "2차비번", user["pw2"]
        )
        frame := SessionProtocol.EncodeFrame("SESSION_ACQUIRE", credentials)
        if !BackgroundProcessManager.SendInput(frame, false) {
            credentials["1차비번"] := ""
            credentials["2차비번"] := ""
            this.Fail("세션 인증정보를 백그라운드 프로세스에 전달하지 못했습니다.")
            return false
        }
        credentials["1차비번"] := ""
        credentials["2차비번"] := ""
        return true
    }

    static HandleWorkerFrame(line) {
        if (SubStr(line, 1, 15) = "SESSION_RESULT|") {
            try {
                frame := SessionProtocol.DecodeFrame(line, "SESSION_RESULT")
                session := SessionStore.FromBundle(frame.Payload)
                if !IsObject(this.CurrentUser) || session.EmployeeId != this.CurrentUser["id"]
                    throw Error("수신 세션의 사용자가 현재 로그인 계정과 다릅니다.")

                this.Session := session
                this.State := "ready"
                this.Save()
                ; 정적 검색에도 구 저장 형식이 소비처처럼 남지 않도록 이름을 분리해 1회 정리만 한다.
                legacyFile := A_ScriptDir "\." "saved_" "cookies" ".json"
                if FileExist(legacyFile)
                    try FileDelete(legacyFile)
                BackgroundProcessManager.SendInput("SESSION_SAVED")
                this.NotifyState("ready")
                LogDebug("통합 CookieJar 세션 준비 완료: 쿠키 " session.CookieJar.Count "개")

                callbacks := this.ReadyCallbacks
                this.ReadyCallbacks := []
                this.ErrorCallbacks := []
                for callback in callbacks {
                    try callback.Call(session)
                    catch as err
                        LogDebug("세션 준비 후 작업 실행 오류: " err.Message)
                }

                ; 작업자 자동 조회보다 근태 조회를 먼저 끝내 업무일지 상태가 누락되지 않게 한다.
                try RequestGuntaeData()
                try RunPendingAutoImportWorkers(session)
                try SetTimer AutoRefreshERPOrder, 3600000
            } catch as err {
                this.Fail(err.Message)
            }
            return
        }

        if (SubStr(line, 1, 14) = "SESSION_ERROR|") {
            try {
                frame := SessionProtocol.DecodeFrame(line, "SESSION_ERROR")
                message := frame.Payload.Has("message") ? frame.Payload["message"] : "통합 로그인 실패"
            } catch as err {
                message := "세션 오류 응답을 해석하지 못했습니다: " err.Message
            }
            this.Fail(message)
        }
    }

    static IsReady(employeeId := "") {
        if (this.State != "ready" || !IsObject(this.Session))
            return false
        if (employeeId != "" && this.Session.EmployeeId != employeeId)
            return false
        return this.Session.CookieJar.Count > 0
    }

    static Request(method, url, body := "", headers := 0) {
        if !this.IsReady()
            throw Error("SESSION_NOT_READY: 통합 로그인 세션이 준비되지 않았습니다.")

        try {
            return this._RequestOnce(method, url, body, headers)
        } catch as firstError {
            if !this.IsExpiredError(firstError)
                throw firstError

            LogDebug("요청에서 세션 만료 감지: 통합 세션을 재획득한 뒤 요청을 1회 재실행합니다.")
            if !this._ReacquireAndWait(&acquireError)
                throw Error("SESSION_RETRY_FAILED: " acquireError)

            try {
                return this._RequestOnce(method, url, body, headers)
            } catch as secondError {
                throw Error("SESSION_RETRY_FAILED: " secondError.Message)
            }
        }
    }

    static _RequestOnce(method, url, body := "", headers := 0) {
        if !this.IsReady()
            throw Error("SESSION_NOT_READY: 통합 로그인 세션이 준비되지 않았습니다.")
        beforeRevision := this.Session.CookieRevision
        try {
            response := this.Session.Request(method, url, body, headers)
            if this._LooksExpiredResponse(response)
                throw Error("SESSION_EXPIRED: 서버가 로그인 만료 응답을 반환했습니다.")
            if (this.Session.CookieRevision != beforeRevision)
                this.Save()
            return response
        } catch as err {
            if this.IsExpiredError(err)
                this.State := "empty"
            throw err
        }
    }

    static _ReacquireAndWait(&errorMessage) {
        waiter := {Done: false, Success: false, Message: ""}
        readyCallback := ObjBindMethod(this, "_CompleteWaitReady", waiter)
        errorCallback := ObjBindMethod(this, "_CompleteWaitError", waiter)
        if !this.Reacquire(readyCallback, errorCallback) {
            errorMessage := "통합 세션 재획득을 시작하지 못했습니다."
            return false
        }

        startedAt := A_TickCount
        while !waiter.Done && (A_TickCount - startedAt < 180000)
            Sleep(50)

        if !waiter.Done {
            BackgroundProcessManager.Cleanup()
            this.Fail("통합 세션 재획득 시간이 초과되었습니다.")
            errorMessage := "통합 세션 재획득 시간이 초과되었습니다."
            return false
        }
        errorMessage := waiter.Message
        return waiter.Success
    }

    static _CompleteWaitReady(waiter, *) {
        waiter.Success := true
        waiter.Done := true
    }

    static _CompleteWaitError(waiter, message) {
        waiter.Message := message
        waiter.Done := true
    }

    static _LooksExpiredResponse(response) {
        lowerUrl := StrLower(response.Url)
        if InStr(lowerUrl, "/user/login.face") || InStr(lowerUrl, "/irj/portal/login")
            return true

        text := response.Text
        if RegExMatch(text, "i)<Parameter[^>]+id=.ErrorCode.[^>]*>\s*-99999")
            return true
        if RegExMatch(text, "i)<title[^>]*>\s*로그\s*아웃\s*</title>")
            return true
        if InStr(lowerUrl, "ep.humetro.busan.kr") {
            lowerText := StrLower(text)
            if InStr(lowerText, 'name="j_password"') || InStr(lowerText, "name='j_password'")
                return true
        }
        return false
    }

    static GetCookieHeaderForUrl(url) {
        if !this.IsReady()
            return ""
        return this.Session.GetCookieHeaderForUrl(url)
    }

    static GetCookiesForCDP() {
        if !this.IsReady()
            return []
        return this.Session.GetCookiesForCDP()
    }

    static HasCookieForUrl(url, name) {
        return this.IsReady() && this.Session.HasCookieForUrl(url, name)
    }

    static Save() {
        if this.IsReady()
            SessionStore.Save(this.Session)
    }

    static Clear(deletePersisted := false) {
        this.Session := 0
        this.State := "empty"
        this.ReadyCallbacks := []
        this.ErrorCallbacks := []
        if deletePersisted
            SessionStore.Clear()
    }

    static Reacquire(onReady := 0, onError := 0) {
        user := IsObject(this.CurrentUser) ? this.CurrentUser : ConfigManager.CurrentUser
        this.Clear(true)
        return this.AcquireAsync(user, onReady, onError)
    }

    static ValidateMis() {
        if !this.IsReady()
            return false
        response := this.Request("POST", SessionAuthenticator.BaseMis "/ssoLogin.do",
            XPlatformProtocol.BuildSsoLoginXml(this.Session.EmployeeId),
            Map("Content-Type", "text/xml;charset=UTF-8"))
        XPlatformProtocol.EnsureSuccess(response.Text, true)
        return true
    }

    static IsExpiredError(err) {
        message := err.Message
        if InStr(message, "SESSION_RETRY_FAILED:") = 1
            return false
        return InStr(message, "SESSION_EXPIRED")
            || InStr(message, "SESSION_NOT_READY")
            || InStr(message, "세션이 만료")
            || InStr(message, "세션이 없")
            || InStr(message, "HTTP 오류 401")
            || InStr(message, "HTTP 오류 403")
    }

    static Fail(message) {
        this.Session := 0
        this.State := "error"
        LogDebug("통합 세션 오류: " message)
        BackgroundProcessManager.SendInput("SESSION_SAVED")
        this.NotifyState("error", message)

        callbacks := this.ErrorCallbacks
        this.ReadyCallbacks := []
        this.ErrorCallbacks := []
        for callback in callbacks {
            try callback.Call(message)
            catch
                continue
        }
    }

    static NotifyState(state, message := "") {
        global wv
        if !wv
            return
        payload := Map("type", "sessionStatus", "state", state)
        if (message != "")
            payload["message"] := message
        try wv.PostWebMessageAsJson(JSON.stringify(payload))
    }
}
