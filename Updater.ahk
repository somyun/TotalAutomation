#Requires AutoHotkey v2.0
#SingleInstance Force
#Include "Lib\JSON.ahk"

; ==============================================================================
; Configuration
; ==============================================================================
configFile := A_ScriptDir "\config.json"
mainExe := A_ScriptDir "\Main.exe"
mainAhk := A_ScriptDir "\Main.ahk"
logFile := A_ScriptDir "\debug.log"

; [New] Asset Configuration
; type: "exe" or "zip"
; asset: Name in GitHub Release
; local: Local destination path (file path for exe, directory path for zip)
Assets := [
    Map("type", "exe", "asset", "Main.exe", "local", mainExe)
]

; ==============================================================================
; GUI Setup (Lazy Init)
; ==============================================================================
CreateGui() {
    g := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner", "Update")
    g.BackColor := "White"
    g.SetFont("s10", "Malgun Gothic")
    g.Add("Text", "w300 Center vStatusText", "업데이트 확인 중...")
    g.Add("Progress", "w300 h20 -Smooth +0x8 vLoadingBar")
    return g
}

global UpdateGui := ""

UpdateStatus(text) {
    global UpdateGui
    if (UpdateGui == "") {
        UpdateGui := CreateGui()
        UpdateGui.Show("NoActivate")
    }
    UpdateGui["StatusText"].Text := text
}

; ==============================================================================
; Main Logic
; ==============================================================================
try {
    ; 0. IMMEDIATE GUI
    UpdateStatus("업데이트 확인 중...")

    ; Safety Mode Check
    mode := (!FileExist(mainExe) && FileExist(mainAhk)) ? "Development" : "Production"
    repo := "MyungjinSong/TotalAutomation"
    currentVer := GetLocalVersion()

    ; 1. INTERNET CHECK
    UpdateStatus("네트워크 연결 확인 중...")
    isOnline := CheckInternetConnection()

    if (!isOnline) {
        ; Case: OFFLINE
        if (FileExist(mainExe)) {
            Run(mainExe " /offline")
        } else if (FileExist(mainAhk)) {
            Run(mainAhk " /offline")
        } else {
            MsgBox("인터넷 연결이 없으며 실행할 파일(Main.exe)도 없습니다.")
        }
        if (UpdateGui != "")
            UpdateGui.Destroy()
        ExitApp
    }

    ; 2. VERSION & INTEGRITY CHECK
    UpdateStatus("버전 정보 가져오는 중...")
    latestRelease := GetLatestRelease(repo)

    shouldUpdate := false
    updateReason := ""
    latestVer := ""
    downloadQueue := [] ; List of maps: {url, path, type, name}

    if (IsObject(latestRelease) && latestRelease.Has("tag_name")) {
        latestVer := latestRelease["tag_name"]

        cVerClean := StrReplace(currentVer, "v", "")
        lVerClean := StrReplace(latestVer, "v", "")

        ; Condition A: New Version Available
        if (VerCompare(lVerClean, cVerClean) > 0) {
            shouldUpdate := true
            updateReason := "새 버전 (" latestVer ")"

            ; In case of update, we download ALL assets
            for item in Assets {
                url := GetAssetUrl(latestRelease, item["asset"])
                if (url != "") {
                    downloadQueue.Push(Map("url", url, "item", item))
                } else {
                    MsgBox("경고: 릴리즈에서 파일을 찾을 수 없습니다: " item["asset"])
                }
            }
        }
        ; Condition B: Missing Files (Repair) - Only if version is matching or we are strictly repairing
        else {
            for item in Assets {
                isMissing := false
                if (item["type"] == "exe" || item["type"] == "file") {
                    if !FileExist(item["local"])
                        isMissing := true
                } else if (item["type"] == "zip") {
                    if !DirExist(item["local"])
                        isMissing := true
                    else {
                        ; Simple check: if directory is empty?
                        ; For now, DirExist is the main check.
                    }
                }

                if (isMissing) {
                    shouldUpdate := true
                    updateReason := "필수 파일 복구"

                    url := GetAssetUrl(latestRelease, item["asset"])
                    if (url != "") {
                        downloadQueue.Push(Map("url", url, "item", item))
                    } else {
                        MsgBox("경고: 복구할 파일을 릴리즈에서 찾을 수 없습니다: " item["asset"])
                    }
                }
            }
        }

    } else {
        ; API Failed or Limit
        MsgBox("업데이트 정보를 받아올 수 없습니다")
        if (UpdateGui != "")
            UpdateGui.Destroy()

        if (mode == "Production" && !ProcessExist("Main.exe") && FileExist(mainExe))
            Run(mainExe " /skipupdate")

        ExitApp
    }

    ; 3. EXECUTE UPDATE
    if (shouldUpdate && downloadQueue.Length > 0 && mode == "Production") {

        UpdateStatus(updateReason " 진행 중...")

        ; Close Main
        if ProcessExist("Main.exe") {
            UpdateStatus("메인 프로그램 종료 중...")
            try ProcessClose("Main.exe")
            if !ProcessWaitClose("Main.exe", 2) {
                try RunWait('taskkill /F /IM "Main.exe"', , "Hide")
                ProcessWaitClose("Main.exe", 1)
            }
        }

        if ProcessExist("Main.exe") {
            MsgBox("Main.exe를 종료할 수 없습니다. 수동으로 종료해 주세요.")
            ExitApp
        }

        UpdateGui["LoadingBar"].Opt("-0x8") ; Determinate
        UpdateGui["LoadingBar"].Range := "0-" downloadQueue.Length * 100
        totalProgress := 0

        ; Process Queue
        for i, task in downloadQueue {
            item := task["item"]
            url := task["url"]
            assetName := item["asset"]

            UpdateStatus("다운로드 중: " assetName)

            tempFile := A_ScriptDir "\temp_" assetName

            try {
                DownloadFile(url, tempFile)

                UpdateStatus("설치 중: " assetName)

                if (item["type"] == "exe" || item["type"] == "file") {
                    ; File Move
                    targetPath := item["local"]
                    if FileExist(targetPath)
                        FileMove(targetPath, targetPath ".old", 1)
                    FileMove(tempFile, targetPath, 1)
                    try FileDelete(targetPath ".old")

                } else if (item["type"] == "zip") {
                    ; Unzip
                    destDir := item["local"]
                    if !DirExist(destDir)
                        DirCreate(destDir)

                    ; Extract to Root (assuming structure in zip is ui/..., Lib/...)
                    UnzipFile(tempFile, A_ScriptDir)

                    FileDelete(tempFile)
                }

                UpdateGui["LoadingBar"].Value := i * 100

            } catch as e {
                MsgBox("설치 실패 (" assetName "): " e.Message)
                ExitApp
            }
        }

        UpdateStatus("업데이트 완료!")
        Sleep 500
        Run(mainExe " /skipupdate")

    } else {
        ; No Update or Dev Mode
        if (mode == "Production" && !ProcessExist("Main.exe")) {
            UpdateStatus("최신 버전입니다.")
            Sleep 500
            Run(mainExe " /skipupdate")
        }
    }

} catch as e {
    MsgBox("오류 발생: " e.Message)
}

if (UpdateGui != "")
    UpdateGui.Destroy()

; ==============================================================================
; Helper Functions
; ==============================================================================

GetLocalVersion() {
    if FileExist(mainExe) {
        try {
            ver := FileGetVersion(mainExe)
            if (ver != "")
                return "v" ver
        }
    }
    if FileExist(mainAhk) {
        content := FileRead(mainAhk, "UTF-8")
        if RegExMatch(content, 'AppVersion\s*:=\s*"(v[\d\.]+)"', &match)
            return match[1]
    }
    return "v0.0.0"
}

GetLatestRelease(repo) {
    url := "https://api.github.com/repos/" repo "/releases/latest"
    try {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        whr.Open("GET", url, true)
        whr.SetRequestHeader("User-Agent", "AutoHotkey")
        whr.Option[4] := 13056
        whr.Send()
        whr.WaitForResponse()
        if (whr.Status == 200)
            return JSON.parse(whr.ResponseText)
    }
    return ""
}

GetAssetUrl(releaseData, assetName) {
    if (releaseData.Has("assets")) {
        for asset in releaseData["assets"] {
            if (asset["name"] = assetName) ; Case-insensitive comparison
                return asset["browser_download_url"]
        }
    }
    return ""
}

CheckInternetConnection() {
    try {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        whr.Open("HEAD", "http://www.google.com", true)
        whr.Option[4] := 13056
        whr.Send()
        whr.WaitForResponse(2)
        return (whr.Status == 200)
    } catch {
        return false
    }
}

DownloadFile(url, dest) {
    whr := ComObject("WinHttp.WinHttpRequest.5.1")
    whr.Open("GET", url, true)
    whr.SetRequestHeader("User-Agent", "AutoHotkey")
    whr.Option[4] := 13056
    whr.Send()
    whr.WaitForResponse()

    if (whr.Status != 200)
        throw Error("Download failed, status: " whr.Status)

    stream := ComObject("ADODB.Stream")
    stream.Type := 1 ; Binary
    stream.Open()
    stream.Write(whr.ResponseBody)
    stream.SaveToFile(dest, 2) ; Overwrite
    stream.Close()
}

UnzipFile(zipPath, destDir) {
    shell := ComObject("Shell.Application")

    ; Ensure dest exists
    if !DirExist(destDir)
        DirCreate(destDir)

    ; Get Zip items
    try {
        zipFolder := shell.NameSpace(zipPath)
        destFolder := shell.NameSpace(destDir)

        if (zipFolder && destFolder) {
            ; 4 (No Progress UI) + 16 (Yes to All)
            destFolder.CopyHere(zipFolder.Items, 20)
            Sleep 1000
        } else {
            throw Error("Failed to open Zip or Dest folder")
        }
    } catch as e {
        throw Error("Unzip failed: " e.Message)
    }
}
