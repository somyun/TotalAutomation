#Requires AutoHotkey v2.0
#SingleInstance Force
#Include "..\Lib\JSON.ahk"
#Include "..\Lib\Chrome.ahk"
#Include "..\Lib\SessionAuth.ahk"

LogDebug(*) {
}

SafeCookieDescription(cookie) {
    location := cookie.Has("url") ? cookie["url"]
        : (cookie.Has("domain") ? cookie["domain"] : "")
    sameSite := cookie.Has("sameSite") ? cookie["sameSite"] : ""
    return "name=" cookie["name"]
        . ", location=" location
        . ", path=" cookie["path"]
        . ", secure=" cookie["secure"]
        . ", sameSite=" sameSite
}

try {
    SessionStore.FilePath := A_ScriptDir "\..\.session.dat"
    session := SessionStore.Load()
    if !IsObject(session)
        throw Error("저장된 통합 세션이 없습니다.")

    chromeInst := Chrome([], , , 9222)
    pages := chromeInst.GetPageList()
    if (pages.Length = 0)
        throw Error("9222 포트에 CDP page가 없습니다.")
    page := Chrome.Page(StrReplace(pages[1]["webSocketDebuggerUrl"], "localhost", "127.0.0.1"))
    page.Call("Network.enable")

    cookies := session.GetCookiesForCDP()
    failures := 0
    for cookie in cookies {
        try {
            result := page.Call("Network.setCookie", cookie)
            if IsObject(result) && result.Has("success") && !result["success"]
                throw Error("Network.setCookie가 success=false를 반환했습니다.")
        } catch as err {
            failures += 1
            FileAppend("CDP COOKIE FAIL: " SafeCookieDescription(cookie)
                . " | " err.Message " | " err.Extra "`n", "*", "UTF-8")
        }
    }

    if (failures > 0)
        throw Error("CDP 쿠키 " failures "개가 거부되었습니다.")

    page.Call("Network.setCookies", Map("cookies", cookies))
    FileAppend("CdpCookieInjectionTest: PASS (" cookies.Length " cookies)`n", "*", "UTF-8")
} catch as err {
    FileAppend("CdpCookieInjectionTest: FAIL line " err.Line " - " err.Message
        . " | " err.Extra "`n", "*", "UTF-8")
    ExitApp(1)
}
ExitApp(0)
