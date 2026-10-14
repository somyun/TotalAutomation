#Requires AutoHotkey v2.0 32-bit
#SingleInstance Off
#NoTrayIcon
#Include %A_ScriptDir%\Lib\JSON.ahk
#Include %A_ScriptDir%\Lib\SapRfcRelease.ahk
#Include %A_ScriptDir%\Lib\SapRfcClient.ahk
#Include %A_ScriptDir%\Lib\SapLocationOrders.ahk
#Include %A_ScriptDir%\Lib\SapMaterialQuery.ahk
SetWorkingDir(A_ScriptDir)
RunSapRfcRequest(request) {
    if !(request is Map) || !request.Has("protocol") || request["protocol"] != SapRfcRelease.Protocol
        throw Error("Invalid protocol")
    command := request["command"]
    if command = "Version"
        return Map("ok", JSON.true, "rows", [])
    if command != "Locations" && command != "LocationMaintenance" && command != "MaterialLocations"
        throw Error("Unsupported command")
    employeeId := request["employeeId"]
    if !RegExMatch(employeeId, "^\d+$")
        throw Error("Invalid employee")
    if command != "MaterialLocations" {
        workCenter := request["workCenter"], year := request["year"]
        if !RegExMatch(workCenter, "^\d+$") || !RegExMatch(year, "^\d{4}$")
            throw Error("Invalid conditions")
        if command = "LocationMaintenance" && !(request["orders"] is Array)
            throw Error("Invalid orders")
    }
    rfc := 0, ticket := "", bundle := 0, matches := []
    try {
        bundle := LoadSession()
        if String(bundle["employeeId"]) != employeeId
            throw Error("User changed")
        for cookie in bundle["cookies"]
            if cookie["name"] = "MYSAPSSO2" && cookie["domain"] = "humetro.busan.kr"
                matches.Push(cookie["value"])
        if matches.Length != 1
            throw Error("Session unavailable")
        ticket := matches[1]
        while RegExMatch(ticket, "%([0-9A-Fa-f]{2})", &match)
            ticket := StrReplace(ticket, match[0], Chr(Integer("0x" match[1])))
        rfc := SapRfcClient()
        rfc.Open(Map("mshost", "erpdb.humetro.busan.kr", "msserv", "3600", "sysid", "BEP", "group", "ERP_BEP", "client", "100", "user", employeeId, "lang", "EN", "mysapsso2", ticket, "trace", "0"))
        if command = "MaterialLocations"
            return Map("ok", JSON.true, "rows", QueryMaterialLocations(rfc, request["locations"], request["search"]))
        if command = "Locations"
            return Map("ok", JSON.true, "rows", QueryLocationOrders(rfc, workCenter, year))
        existingRows := QueryExistingLocationOrders(rfc, workCenter, request["orders"]), needsYearOrders := false
        for row in existingRows {
            if (row["ORDER_TYPE"] = "PM01" || row["ORDER_TYPE"] = "PM03")
                && RegExMatch(row["START_DATE"], "^\d{8}$") && row["START_DATE"] != "00000000"
                && SubStr(row["START_DATE"], 1, 4) < year
                needsYearOrders := true
        }
        currentRows := needsYearOrders ? QueryLocationOrders(rfc, workCenter, year) : []
        return Map("ok", JSON.true, "rows", [], "existingRows", existingRows, "currentRows", currentRows, "year", year)
    } finally {
        if IsObject(rfc)
            rfc.Close()
        ticket := "", bundle := 0, matches := []
    }
}
outputPath := "", exitCode := 1
try {
    if A_Args.Length != 2
        throw Error("Invalid arguments")
    outputPath := A_Args[2]
    result := RunSapRfcRequest(JSON.parse(FileRead(A_Args[1], "UTF-8")))
    exitCode := 0
} catch {
    result := Map("ok", JSON.false, "error", "SAP RFC 조회에 실패했습니다. 로그인 세션·접속 환경·권한을 확인해주세요.")
}
if outputPath != "" {
    try {
        result["protocol"] := SapRfcRelease.Protocol, result["version"] := SapRfcRelease.Version
        outFile := FileOpen(outputPath, "w", "UTF-8-RAW")
        if !outFile
            ExitApp(1)
        outFile.Write(JSON.stringify(result)), outFile.Close()
    } catch {
        ExitApp(1)
    }
}
ExitApp(exitCode)
