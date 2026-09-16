#Requires AutoHotkey v2.0
#SingleInstance Ignore
#Include "Lib\JSON.ahk"
#Include "Lib\SessionAuth.ahk"

global DEBUG_FILE := A_ScriptDir "\debug_bg.txt"

LogDebug(message) {
    global DEBUG_FILE
    try {
        safe := StrReplace(StrReplace(message, "`r", " "), "`n", " ")
        safe := RegExReplace(safe,
            "i)(password|secondpwd|pwd|cookie|set-cookie|mysapsso2|jsessionid|\bK)\s*[=:]\s*[^\s,;&]+",
            "$1=[REDACTED]")
        FileAppend(FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") " [SessionBG] " safe "`n",
            DEBUG_FILE, "UTF-8")
    }
}

WriteFrame(kind, payload) {
    FileAppend(SessionProtocol.EncodeFrame(kind, payload), "*", "UTF-8")
}

ReadAcquireRequest() {
    ; 소스 파일을 직접 실행한 경우에만 로컬 디버그 자격정보를 사용한다.
    ; Main이 소스 worker를 실행할 때는 --session-worker 표식과 stdin 프레임을 사용한다.
    if !A_IsCompiled && (A_Args.Length = 0 || A_Args[1] != "--session-worker") {
        debugCredentialFile := A_ScriptDir "\세션BG.txt"
        if !FileExist(debugCredentialFile)
            throw Error("단독 디버그 실행용 세션BG.txt 파일이 없습니다.")
        parts := StrSplit(Trim(FileRead(debugCredentialFile, "UTF-8"), "`r`n "), "|")
        if (parts.Length < 3)
            throw Error("세션BG.txt 형식은 사번|1차비번|2차비번 이어야 합니다.")
        LogDebug("단독 디버그 모드: 로컬 인증정보 파일을 읽었습니다. 사용자=" Trim(parts[1]))
        return {
            Credentials: Map(
                "사번", Trim(parts[1]),
                "1차비번", Trim(parts[2]),
                "2차비번", Trim(parts[3])
            ),
            Input: 0,
            StandaloneDebug: true,
            ShowDialog: !(A_Args.Length > 0 && A_Args[1] = "--debug-no-ui")
        }
    }

    stdin := FileOpen("*", "r", "UTF-8")
    line := Trim(stdin.ReadLine(), "`r`n ")
    if (line = "")
        throw Error("Main 프로세스에서 인증 요청을 받지 못했습니다.")
    frame := SessionProtocol.DecodeFrame(line, "SESSION_ACQUIRE")
    return {Credentials: frame.Payload, Input: stdin, StandaloneDebug: false}
}

WaitForSavedAck(stdin) {
    try {
        SetTimer(() => ExitApp(), -10000)
        loop {
            line := Trim(stdin.ReadLine(), "`r`n ")
            if (line = "SESSION_SAVED" || line = "")
                break
        }
        SetTimer(() => ExitApp(), 0)
    }
}

LogDebug("통합 세션 백그라운드 획득 시작")

try {
    request := ReadAcquireRequest()
    credentials := request.Credentials
    for key in ["사번", "1차비번", "2차비번"] {
        if !credentials.Has(key) || credentials[key] = ""
            throw Error("통합 로그인 인증정보가 완전하지 않습니다.")
    }

    employeeId := credentials["사번"]
    session := SessionAuthenticator.Acquire(credentials)
    credentials["1차비번"] := ""
    credentials["2차비번"] := ""

    if (session.EmployeeId != employeeId)
        throw Error("획득한 세션의 사용자가 요청 계정과 일치하지 않습니다.")

    SessionStore.Save(session)
    LogDebug("통합 세션 획득 완료: CookieJar " session.CookieJar.Count "개")
    if request.StandaloneDebug {
        if request.ShowDialog
            MsgBox("통합 세션 확보에 성공했습니다.`nCookieJar: " session.CookieJar.Count "개", "세션BG 디버그")
    } else {
        WriteFrame("SESSION_RESULT", SessionStore.BuildBundle(session))
        WaitForSavedAck(request.Input)
    }
    ExitApp(0)
} catch as err {
    try {
        if IsSet(credentials) && IsObject(credentials) {
            if credentials.Has("1차비번")
                credentials["1차비번"] := ""
            if credentials.Has("2차비번")
                credentials["2차비번"] := ""
        }
    }
    detail := err.Message " | file=" err.File " | line=" err.Line " | what=" err.What
    LogDebug("통합 세션 획득 실패: " detail)
    if IsSet(request) && request.HasOwnProp("StandaloneDebug") && request.StandaloneDebug
        && request.ShowDialog
        MsgBox("통합 세션 확보에 실패했습니다.`n" err.Message "`nLine: " err.Line, "세션BG 디버그", "Iconx")
    else
        try WriteFrame("SESSION_ERROR", Map("message", err.Message, "line", err.Line))
    ExitApp(1)
}
