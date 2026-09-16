#Requires AutoHotkey v2.0
#SingleInstance Force
#Include "..\Lib\JSON.ahk"
#Include "..\Lib\SessionAuth.ahk"

class SessionManager {
    static Request(*) {
        throw Error("네트워크 요청은 오프라인 단위 테스트 범위가 아닙니다.")
    }
}

#Include "..\승인정보조회.ahk"

Assert(condition, message) {
    if !condition
        throw Error("ASSERT FAILED: " message)
}

LogDebug(*) {
}

try {
session := HttpSession()
session.EmployeeId := "123456"
session.StoreCookie(session.ParseUrl("https://btcep.humetro.busan.kr/portal/"),
    "JSESSIONID=PORTAL; Path=/; Secure; HttpOnly")
session.StoreCookie(session.ParseUrl("https://sso.humetro.busan.kr/sso/pmi-sso2.jsp"),
    "JSESSIONID=SSO; Path=/sso; Secure")
session.StoreCookie(session.ParseUrl("https://mis.humetro.busan.kr/FS/index.jsp"),
    "JSESSIONID=MISROOT; Path=/FS; Secure")
session.StoreCookie(session.ParseUrl("https://mis.humetro.busan.kr/FS/private/check"),
    "JSESSIONID=MISPRIVATE; Path=/FS/private; Secure")
session.StoreCookie(session.ParseUrl("http://ep.humetro.busan.kr/irj/portal"),
    "MYSAPSSO2=SAP; Domain=humetro.busan.kr; Path=/; HttpOnly")

Assert(session.CookieJar.Count = 5, "동일 이름의 도메인/경로별 쿠키가 공존해야 함")
Assert(session.HasCookieForUrl("https://mis.humetro.busan.kr/FS/selectList.do", "JSESSIONID"),
    "MIS /FS 쿠키 선택")
misHeader := session.GetCookieHeaderForUrl("https://mis.humetro.busan.kr/FS/private/check")
Assert(InStr(misHeader, "JSESSIONID=MISPRIVATE") = 1, "긴 Path 쿠키가 먼저 정렬되어야 함")
Assert(InStr(misHeader, "JSESSIONID=MISROOT"), "상위 Path 쿠키도 포함되어야 함")
Assert(InStr(misHeader, "MYSAPSSO2=SAP"), "공통 Domain 쿠키가 포함되어야 함")
Assert(!InStr(session.GetCookieHeaderForUrl("http://mis.humetro.busan.kr/FS/selectList.do"), "MISROOT"),
    "Secure 쿠키를 HTTP로 보내면 안 됨")
beforeRejectedCookie := session.CookieJar.Count
session.StoreCookie(session.ParseUrl("https://mis.humetro.busan.kr/FS/index.jsp"),
    "OUTSIDE=NO; Domain=busan.kr; Path=/")
Assert(session.CookieJar.Count = beforeRejectedCookie, "허용 범위 밖 Domain 쿠키를 거부해야 함")
session.StoreCookie(session.ParseUrl("https://mis.humetro.busan.kr/FS/index.jsp"),
    "EXPIRED=NO; Path=/FS; Expires=Thu, 01 Jan 1970 00:00:00 GMT")
Assert(!session.HasCookieForUrl("https://mis.humetro.busan.kr/FS/index.jsp", "EXPIRED"),
    "만료 쿠키를 전송하면 안 됨")

exported := session.ExportCookies()
exportJson := JSON.stringify(exported)
Assert(InStr(exportJson, '"secure":true') && InStr(exportJson, '"hostOnly":true'),
    "세션 번들의 쿠키 속성은 JSON Boolean이어야 함")
restored := HttpSession()
restored.EmployeeId := session.EmployeeId
restored.ImportCookies(exported)
Assert(restored.CookieJar.Count = session.CookieJar.Count, "CookieJar 직렬화 왕복")
Assert(restored.GetCookieHeaderForUrl("https://mis.humetro.busan.kr/FS/private/check") = misHeader,
    "복원 후 Cookie 헤더가 동일해야 함")
Assert(restored.GetCookiesForCDP().Length = session.CookieJar.Count, "CDP 쿠키 변환 개수")
cdpJson := JSON.stringify(restored.GetCookiesForCDP())
Assert(InStr(cdpJson, '"secure":true'), "CDP secure 속성은 JSON Boolean이어야 함")
Assert(InStr(cdpJson, '"httpOnly":false'), "CDP httpOnly 속성은 JSON Boolean이어야 함")
Assert(!InStr(cdpJson, '"secure":1') && !InStr(cdpJson, '"httpOnly":0'),
    "CDP Boolean 속성을 숫자로 직렬화하면 안 됨")

largeBundle := SessionStore.BuildBundle(restored)
Loop 150
    largeBundle["cookies"].Push(exported[1])
frame := SessionProtocol.EncodeFrame("SESSION_RESULT", largeBundle)
Assert(StrLen(frame) > 8192, "분할 시험 프레임이 8KB보다 커야 함")
decoded := SessionProtocol.DecodeFrame(Trim(frame, "`r`n"), "SESSION_RESULT")
Assert(decoded.Payload["cookies"].Length = largeBundle["cookies"].Length,
    "8KB 이상 프레임 Base64 왕복")

pipeBuffer := ""
parsedFrames := []
offset := 1
chunkSizes := [8096, 37, 4093, 211]
chunkIndex := 1
while (offset <= StrLen(frame)) {
    chunkSize := chunkSizes[chunkIndex]
    chunkIndex := chunkIndex = chunkSizes.Length ? 1 : chunkIndex + 1
    extracted := SessionProtocol.ExtractFrames(pipeBuffer, SubStr(frame, offset, chunkSize))
    pipeBuffer := extracted.Buffer
    for completed in extracted.Frames
        parsedFrames.Push(completed)
    offset += chunkSize
}
Assert(parsedFrames.Length = 1, "분할된 파이프 프레임을 정확히 한 번 완성해야 함")
Assert(pipeBuffer = "", "완성 프레임 뒤 누적 버퍼가 비어야 함")
splitDecoded := SessionProtocol.DecodeFrame(parsedFrames[1], "SESSION_RESULT")
Assert(splitDecoded.Payload["cookies"].Length = largeBundle["cookies"].Length,
    "분할 프레임 재조립 후 데이터가 동일해야 함")

approvalXml := '<?xml version="1.0"?><Root><Parameters><Parameter id="ErrorCode">0</Parameter>'
    . '</Parameters><Dataset id="ds_out"><Rows>'
    . '<Row><Col id="DRIV_NAME"> 송 명 진 </Col><Col id="WORK_NAME">철도장비운행 승인</Col>'
    . '<Col id="APPR_NUMB">FIRST</Col><Col id="CNAP_BSNM">첫부서</Col><Col id="CNAP_NAME">첫승인자</Col></Row>'
    . '<Row><Col id="DRIV_NAME">송명진</Col><Col id="WORK_NAME">철도장비운행</Col>'
    . '<Col id="APPR_NUMB">SECOND</Col><Col id="CNAP_BSNM">둘째부서</Col><Col id="CNAP_NAME">둘째승인자</Col></Row>'
    . '</Rows></Dataset></Root>'
approvalDoc := XPlatformProtocol.EnsureSuccess(approvalXml)
approvalSelection := ApprovalInfoService.SelectFirst(approvalDoc,
    ApprovalInfoService.NormalizeDriverName("송명진"))
Assert(approvalSelection.Count = 2, "승인정보 전체 일치 건수")
Assert(approvalSelection.First["ApprovalNo"] = "FIRST", "서버 응답의 첫 승인정보를 선택해야 함")
Assert(approvalSelection.First["Dept"] = "첫부서", "CNAP_BSNM 승인부서 매핑")
Assert(approvalSelection.First["Approver"] = "첫승인자", "CNAP_NAME 승인자 매핑")

originalPath := SessionStore.FilePath
SessionStore.FilePath := A_Temp "\session_core_test_" A_TickCount ".dat"
try {
    SessionStore.Save(restored)
    loaded := SessionStore.Load("123456")
    Assert(IsObject(loaded), "DPAPI 세션 복원")
    Assert(loaded.EmployeeId = "123456", "DPAPI 복원 계정")
    Assert(loaded.CookieJar.Count = restored.CookieJar.Count, "DPAPI CookieJar 왕복")
} finally {
    SessionStore.Clear()
    SessionStore.FilePath := originalPath
}

FileAppend("SessionCoreTest: PASS`n", "*", "UTF-8")
} catch as err {
    FileAppend("SessionCoreTest: FAIL line " err.Line " - " err.Message "`n", "*", "UTF-8")
    ExitApp(1)
}
ExitApp(0)
