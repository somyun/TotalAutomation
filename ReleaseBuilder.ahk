#Requires AutoHotkey v2.0 64-bit
#SingleInstance Force
#NoTrayIcon
;@Ahk2Exe-SetMainIcon ui\assets\icon.ico
;@Ahk2Exe-SetDescription 통합자동화 릴리스 빌더
;@Ahk2Exe-SetVersion 1.0.0.0
;@Ahk2Exe-SetProductVersion 1.0.0

SetWorkingDir(A_ScriptDir)

global BuilderBusy := true
global BuilderAutoExit := A_Args.Length = 1 && A_Args[1] = "--build-and-exit"
global BuilderLogPath := A_ScriptDir "\release-builder.log"
global BuilderGui := Gui("+AlwaysOnTop", "통합자동화 릴리스 빌더")
BuilderGui.SetFont("s10", "Malgun Gothic")
BuilderGui.AddText("w560", "SAP RFC 포함 배포용 Main.exe를 생성합니다.")
global BuilderStatus := BuilderGui.AddText("w560", "빌드 준비 중...")
global BuilderProgress := BuilderGui.AddProgress("w560 h20 Range0-100", 0)
global BuilderLog := BuilderGui.AddEdit("w560 r14 ReadOnly -Wrap")
global BuilderClose := BuilderGui.AddButton("w100 Default Disabled", "닫기")
BuilderClose.OnEvent("Click", (*) => ExitApp())
BuilderGui.OnEvent("Close", OnBuilderClose)
if !BuilderAutoExit
    BuilderGui.Show()
SetTimer(StartReleaseBuild, -100)

OnBuilderClose(*) {
    global BuilderBusy
    if !BuilderBusy
        ExitApp()
}

StartReleaseBuild() {
    global BuilderBusy, BuilderClose, BuilderLogPath, BuilderAutoExit
    succeeded := false
    try {
        try FileDelete(BuilderLogPath)
        AppendBuildLog("빌드를 시작합니다.")
        BuildRelease()
        SetBuildStatus("완료: Main.exe가 생성되고 검증되었습니다.", 100)
        AppendBuildLog("빌드 및 패키지 검증이 완료되었습니다.")
        succeeded := true
        if !BuilderAutoExit
            SetBuildStatus("Main.exe 빌드가 완료되었습니다.`n`n닫기를 눌러 종료하세요.",
                100)
    } catch as err {
        SetBuildStatus("실패: " err.Message, 100)
        AppendBuildLog("오류: " err.Message)
        if !BuilderAutoExit
            MsgBox("빌드에 실패했습니다.`n`n" err.Message "`n`n상세 기록: release-builder.log",
                "빌드 실패", "Iconx")
    } finally {
        BuilderBusy := false
        BuilderClose.Enabled := true
    }
    if BuilderAutoExit
        ExitApp(succeeded ? 0 : 1)
}

BuildRelease() {
    compiler := "C:\Program Files\AutoHotkey\Compiler\Ahk2Exe.exe"
    base32 := "C:\Program Files\AutoHotkey\v2\AutoHotkey32.exe"
    base64 := "C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe"
    iconPath := A_ScriptDir "\ui\assets\icon.ico"
    sapSource := A_ScriptDir "\SAPRFC.ahk"
    mainSource := A_ScriptDir "\Main.ahk"
    sapBuild := A_ScriptDir "\SAPRFC.build.exe"
    sapRelease := A_ScriptDir "\SAPRFC.exe"
    mainBuild := A_ScriptDir "\Main.release.exe"
    mainRelease := A_ScriptDir "\Main.exe"

    SetBuildStatus("필수 파일 확인 중...", 5)
    for requiredPath in [compiler, base32, base64, iconPath, sapSource, mainSource] {
        if !FileExist(requiredPath)
            throw Error("필수 파일을 찾을 수 없습니다: " requiredPath)
    }
    if ProcessExist("SAPRFC.exe")
        throw Error("SAPRFC.exe가 실행 중입니다. SAP 조회를 종료한 뒤 다시 실행해주세요.")
    if ProcessExist("Main.exe")
        throw Error("Main.exe가 실행 중입니다. 프로그램을 종료한 뒤 다시 실행해주세요.")

    SetBuildStatus("1/5 SAPRFC.exe 32비트 빌드 중...", 15)
    DeleteIfExists(sapBuild)
    CompileAhk(compiler, sapSource, sapBuild, base32)
    AssertPeMachine(sapBuild, 0x14C, "SAPRFC.build.exe")
    FileCopy(sapBuild, sapRelease, true)
    AppendBuildLog("SAPRFC.exe x86 빌드 완료")

    SetBuildStatus("2/5 SAP RFC 프로토콜 확인 중...", 35)
    releaseVersion := VerifySapRfcVersion(sapRelease)
    AppendBuildLog("SAP RFC 프로토콜 1 / 릴리스 " releaseVersion " 확인 완료")

    SetBuildStatus("3/5 Main.exe 64비트 빌드 중...", 50)
    DeleteIfExists(mainBuild)
    CompileAhk(compiler, mainSource, mainBuild, base64, iconPath)
    AssertPeMachine(mainBuild, 0x8664, "Main.release.exe")
    AppendBuildLog("Main.release.exe x64 빌드 및 아이콘 적용 완료")

    SetBuildStatus("4/5 내장 SAPRFC.exe 추출 검사 중...", 70)
    VerifyEmbeddedPackage(mainBuild, sapRelease)
    AppendBuildLog("Main.exe 내장 SAPRFC.exe 추출 및 바이트 일치 확인 완료")

    SetBuildStatus("5/5 최종 Main.exe 교체 및 검증 중...", 90)
    FileCopy(mainBuild, mainRelease, true)
    if !FilesEqual(mainBuild, mainRelease)
        throw Error("최종 Main.exe 복사 검증에 실패했습니다.")
    AssertPeMachine(mainRelease, 0x8664, "Main.exe")

    DeleteIfExists(sapBuild)
    DeleteIfExists(mainBuild)
    AppendBuildLog("최종 산출물: " mainRelease " (SAP RFC 릴리스 " releaseVersion ")")
}

CompileAhk(compiler, source, output, base, iconPath := "") {
    command := QuoteArg(compiler)
        . " /in " QuoteArg(source)
        . " /out " QuoteArg(output)
        . " /base " QuoteArg(base)
        . " /silent verbose"
    if iconPath != ""
        command .= " /icon " QuoteArg(iconPath)

    exitCode := RunWithTimeout(command, A_ScriptDir, 60000)
    if exitCode != 0 || !FileExist(output)
        throw Error("컴파일에 실패했습니다: " source " (종료코드 " exitCode ")")
}

VerifySapRfcVersion(sapRelease) {
    token := DllCall("GetCurrentProcessId", "UInt") "-" A_TickCount
    requestPath := A_Temp "\sap-release-version-" token ".json"
    responsePath := requestPath ".result"
    try {
        WriteUtf8Raw(requestPath, '{"protocol":1,"command":"Version"}')
        command := QuoteArg(sapRelease) " " QuoteArg(requestPath) " " QuoteArg(responsePath)
        exitCode := RunWithTimeout(command, A_ScriptDir, 10000)
        WaitForProcessNameClose("SAPRFC.exe", 2000)
        if exitCode != 0 || !FileExist(responsePath)
            throw Error("SAP RFC 버전 확인에 실패했습니다. 종료코드: " exitCode)

        response := FileRead(responsePath, "UTF-8")
        if !RegExMatch(response, 'i)"ok"\s*:\s*true')
            || !RegExMatch(response, 'i)"protocol"\s*:\s*1(?:\D|$)')
            || !RegExMatch(response, 'i)"version"\s*:\s*"([^"]+)"', &versionMatch)
            throw Error("SAP RFC 버전 응답 형식이 올바르지 않습니다.")
        return versionMatch[1]
    } finally {
        DeleteIfExists(requestPath)
        DeleteIfExists(responsePath)
    }
}

VerifyEmbeddedPackage(mainBuild, sapRelease) {
    token := DllCall("GetCurrentProcessId", "UInt") "-" A_TickCount
    packageFolder := A_Temp "\sap-package-" token
    packageMain := packageFolder "\Main.exe"
    packageSap := packageFolder "\SAPRFC.exe"
    packageVerified := false
    try {
        WaitForProcessNameClose("SAPRFC.exe", 2000)
        DirCreate(packageFolder)
        FileCopy(mainBuild, packageMain)
        lastError := ""
        loop 3 {
            attempt := A_Index
            try {
                exitCode := RunWithTimeout(QuoteArg(packageMain) " --verify-sap-package", packageFolder, 30000)
                if exitCode != 0 || !FileExist(packageSap)
                    throw Error("종료코드 " exitCode)
                lastError := ""
                break
            } catch as err {
                lastError := err.Message
                if attempt < 3 {
                    AppendBuildLog("내장 패키지 검사 " attempt "차 실패, 다시 시도합니다.")
                    WaitForProcessNameClose("SAPRFC.exe", 3000)
                    Sleep(1000)
                }
            }
        }
        if lastError != ""
            throw Error("Main.exe 내장 SAPRFC.exe 추출 검사에 실패했습니다: " lastError)
        AssertPeMachine(packageSap, 0x14C, "추출된 SAPRFC.exe")
        if !FilesEqual(sapRelease, packageSap)
            throw Error("Main.exe에서 추출한 SAPRFC.exe가 빌드 원본과 다릅니다.")
        packageVerified := true
    } finally {
        if packageVerified && IsSafePackageFolder(packageFolder) && DirExist(packageFolder) {
            try DirDelete(packageFolder, true)
        } else if DirExist(packageFolder) {
            AppendBuildLog("실패 진단용 검사 폴더를 보존했습니다: " packageFolder)
        }
    }
}

WaitForProcessNameClose(processName, timeoutMs) {
    started := A_TickCount
    while ProcessExist(processName) {
        if A_TickCount - started >= timeoutMs
            throw Error(processName " 프로세스가 종료되지 않았습니다.")
        Sleep(50)
    }
}

RunWithTimeout(command, workingDirectory, timeoutMs) {
    Run(command, workingDirectory, "Hide", &pid)
    processHandle := DllCall("OpenProcess", "UInt", 0x101000, "Int", false, "UInt", pid, "Ptr")
    if !processHandle {
        try ProcessClose(pid)
        throw Error("실행한 프로세스의 핸들을 열 수 없습니다: " pid)
    }
    try {
        started := A_TickCount
        loop {
            waitResult := DllCall("WaitForSingleObject", "Ptr", processHandle, "UInt", 50, "UInt")
            if waitResult = 0
                break
            if waitResult != 0x102
                throw Error("프로세스 대기 중 오류가 발생했습니다: " waitResult)
            if A_TickCount - started >= timeoutMs {
                try ProcessClose(pid)
                DllCall("WaitForSingleObject", "Ptr", processHandle, "UInt", 5000, "UInt")
                throw Error("작업 시간이 초과되었습니다: " command)
            }
            Sleep(10)
        }
        exitCode := 0
        if !DllCall("GetExitCodeProcess", "Ptr", processHandle, "UIntP", &exitCode, "Int")
            throw Error("프로세스 종료코드를 확인할 수 없습니다.")
        return exitCode
    } finally {
        DllCall("CloseHandle", "Ptr", processHandle)
    }
}

AssertPeMachine(path, expectedMachine, label) {
    file := FileOpen(path, "r")
    if !file
        throw Error(label " 파일을 열 수 없습니다.")
    try {
        header := Buffer(64, 0)
        if file.RawRead(header, 64) != 64 || NumGet(header, 0, "UShort") != 0x5A4D
            throw Error(label " 파일이 올바른 PE 실행 파일이 아닙니다.")
        peOffset := NumGet(header, 60, "UInt")
        file.Pos := peOffset
        peHeader := Buffer(6, 0)
        if file.RawRead(peHeader, 6) != 6 || NumGet(peHeader, 0, "UInt") != 0x4550
            throw Error(label " PE 헤더를 확인할 수 없습니다.")
        actualMachine := NumGet(peHeader, 4, "UShort")
        if actualMachine != expectedMachine
            throw Error(label " 아키텍처가 올바르지 않습니다. 예상=" Format("0x{:X}", expectedMachine)
                ", 실제=" Format("0x{:X}", actualMachine))
    } finally {
        file.Close()
    }
}

FilesEqual(firstPath, secondPath) {
    if !FileExist(firstPath) || !FileExist(secondPath)
        return false
    if FileGetSize(firstPath) != FileGetSize(secondPath)
        return false

    first := FileOpen(firstPath, "r"), second := FileOpen(secondPath, "r")
    if !first || !second
        return false
    try {
        firstBuffer := Buffer(65536), secondBuffer := Buffer(65536)
        loop {
            firstRead := first.RawRead(firstBuffer, firstBuffer.Size)
            secondRead := second.RawRead(secondBuffer, secondBuffer.Size)
            if firstRead != secondRead
                return false
            if firstRead = 0
                return true
            if DllCall("msvcrt\memcmp", "Ptr", firstBuffer.Ptr, "Ptr", secondBuffer.Ptr,
                "UPtr", firstRead, "Int") != 0
                return false
        }
    } finally {
        first.Close(), second.Close()
    }
}

WriteUtf8Raw(path, text) {
    file := FileOpen(path, "w", "UTF-8-RAW")
    if !file
        throw Error("임시 요청 파일을 만들 수 없습니다.")
    try file.Write(text)
    finally file.Close()
}

QuoteArg(value) {
    return '"' StrReplace(value, '"', '\"') '"'
}

DeleteIfExists(path) {
    if FileExist(path)
        FileDelete(path)
}

IsSafePackageFolder(path) {
    tempPrefix := RTrim(A_Temp, "\/") "\sap-package-"
    return InStr(StrLower(path), StrLower(tempPrefix)) = 1
}

SetBuildStatus(text, progress) {
    global BuilderStatus, BuilderProgress
    BuilderStatus.Text := text
    BuilderProgress.Value := progress
    AppendBuildLog(text)
    Sleep(20)
}

AppendBuildLog(text) {
    global BuilderLog, BuilderLogPath
    line := FormatTime(, "HH:mm:ss") "  " text
    BuilderLog.Value .= (BuilderLog.Value = "" ? "" : "`r`n") line
    SendMessage(0x115, 7, 0, BuilderLog.Hwnd)
    try FileAppend(line "`r`n", BuilderLogPath, "UTF-8")
}
