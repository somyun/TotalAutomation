class SapRfcService {
    static Ready := !A_IsCompiled
    static Busy := false
    static VerifyHelper() => this.Execute(Map("command", "Version"), 10000)
    static Query(request) {
        if !this.Ready
            throw Error("SAPRFC.exe 설치가 완료되지 않았습니다. 프로그램을 다시 실행해주세요.")
        this.VerifyHelper()
        SessionManager.Save()
        return this.Execute(request)
    }
    static Execute(request, timeoutMs := 180000) {
        if this.Busy
            throw Error("SAP 조회가 진행 중입니다. 완료 후 다시 시도해주세요.")
        helper := A_ScriptDir "\SAPRFC.exe"
        if !FileExist(helper)
            throw Error("SAPRFC.exe가 없습니다. 프로그램을 다시 실행해주세요.")
        this.Busy := true, pid := 0, processHandle := 0
        outputPath := A_Temp "\sap-rfc-" DllCall("GetCurrentProcessId", "UInt") "-" A_TickCount ".json", inputPath := outputPath ".request"
        try {
            request["protocol"] := SapRfcRelease.Protocol
            outFile := FileOpen(inputPath, "w", "UTF-8-RAW")
            outFile.Write(JSON.stringify(request)), outFile.Close()
            Run('"' helper '" "' inputPath '" "' outputPath '"', A_ScriptDir, "Hide", &pid)
            processHandle := DllCall("OpenProcess", "UInt", 0x101001, "Int", false, "UInt", pid, "Ptr")
            if !processHandle
                throw Error("SAP 조회 프로세스를 확인할 수 없습니다.")
            started := A_TickCount
            Loop {
                waitResult := DllCall("WaitForSingleObject", "Ptr", processHandle, "UInt", 0, "UInt")
                if waitResult = 0
                    break
                if waitResult != 258
                    throw Error("SAP 조회 프로세스 대기에 실패했습니다.")
                if A_TickCount - started > timeoutMs
                    throw Error("SAP 조회 시간이 초과되었습니다. 잠시 후 다시 시도해주세요.")
                Sleep(100)
            }
            exitCode := 0
            if !DllCall("GetExitCodeProcess", "Ptr", processHandle, "UIntP", &exitCode, "Int")
                throw Error("SAP 조회 종료코드를 확인할 수 없습니다.")
            pid := 0
            if !FileExist(outputPath)
                throw Error("SAP 조회 결과를 받지 못했습니다.")
            try result := JSON.parse(FileRead(outputPath, "UTF-8"))
            catch
                throw Error("SAP 조회 결과 형식이 올바르지 않습니다.")
            if !(result is Map) || !result.Has("protocol") || !result.Has("version")
                || result["protocol"] != SapRfcRelease.Protocol || result["version"] != SapRfcRelease.Version
                throw Error("SAPRFC.exe 버전이 맞지 않습니다. 프로그램을 다시 실행해주세요.")
            if exitCode != 0 || !result.Has("ok") || !result["ok"]
                throw Error("SAP RFC 조회에 실패했습니다. 로그인 세션·접속 환경·권한을 확인해주세요.")
            if !result.Has("rows") || !(result["rows"] is Array)
                throw Error("SAP 조회 결과 형식이 올바르지 않습니다.")
            return result
        } finally {
            if pid {
                if processHandle
                    DllCall("TerminateProcess", "Ptr", processHandle, "UInt", 1)
                else if ProcessExist(pid)
                    try ProcessClose(pid)
            }
            if processHandle
                DllCall("CloseHandle", "Ptr", processHandle)
            for path in [inputPath, outputPath]
                if FileExist(path)
                    try FileDelete(path)
            this.Busy := false
        }
    }
}
