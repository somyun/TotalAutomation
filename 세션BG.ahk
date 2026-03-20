#Requires AutoHotkey v2
#Include Lib\JSON.ahk
#Include Lib\chrome.ahk

; ==============================================================================
; [설정] 초기화 및 로깅 시작
; ==============================================================================
StartTime := A_TickCount
DEBUG_FILE := A_ScriptDir "\debug_bg.txt"
FileAppend("`n", DEBUG_FILE, "UTF-8")

LogDebug("========== 세션BG 시작 ==========")

if (!A_IsCompiled) {
    ; [디버깅 모드] 컴파일되지 않은 경우 (스크립트 실행)
    debugFile := A_ScriptDir "\세션bg.txt"
    if FileExist(debugFile) {
        LogDebug("[설정] 디버깅 모드 진입: " debugFile " 읽기 시도")
        try {
            content := FileRead(debugFile, "UTF-8")
            parts := StrSplit(content, "|")
            if (parts.Length >= 2) {
                sUser := Trim(parts[1])
                sPw1 := Trim(parts[2])
                sPw2 := Trim(parts[3])
                LogDebug("[설정] 디버깅 모드 계정 획득: ID=" sUser ", PW1=" RegExReplace(sPw1, ".", "*", , , 2) ", PW2=" SubStr(
                    sPw2, 1, 1) "*****")
            } else {
                LogDebug("[설정] 오류: 디버깅 파일(세션bg.txt) 형식이 '아이디|비밀번호1|비밀번호2' 여야 합니다.")
                ExitApp(1)
            }
        } catch as e {
            LogDebug("[설정] 오류: 디버깅 파일 읽기 실패: " e.Message)
            ExitApp(1)
        }
    } else {
        LogDebug("[설정] 오류: 디버깅 파일 없음: " debugFile)
        FileAppend('{"error": "Debug file not found"}', "*")
        ExitApp(1)
    }
} else {
    ; [운영 모드] 컴파일된 실행 파일
    LogDebug("[설정] 운영 모드 진입 (Main.ahk에서 세션BG.exe 호출)")
    if (A_Args.Length < 3) {
        LogDebug("[설정] 오류: 인자 부족 (" A_Args.Length "개)")
        FileAppend('{"error": "Invalid arguments"}', "*")
        ExitApp(1)
    }
    sUser := A_Args[1]
    sPw1 := A_Args[2]
    sPw2 := A_Args[3]
}

LogDebug("[설정] 인증 대상 사용자 ID: " sUser)

; ==============================================================================
; [설정] WinHttp 및 쿠키 저장소
; ==============================================================================
http := ComObject("WinHttp.WinHttpRequest.5.1")
http.Option[3] := 0  ; EnableCookieExchange = False
http.Option[6] := 0  ; EnableRedirects = False
http.Option[4] := 13056 ; IgnoreCertErrors

; 도메인별 쿠키 바구니
btcep_cookies := Map()
niw_cookies := Map()
ep_cookies := Map()

; 현재 타겟 바구니
current_cookies := btcep_cookies

; [함수] 쿠키 업데이트
UpdateCookies(headers) {
    global current_cookies
    loop parse, headers, "`n", "`r" {
        if RegExMatch(A_LoopField, "i)^Set-Cookie:\s*([^=]+)=([^;]+)", &m)
            current_cookies[m[1]] := m[2]
    }
}

; [함수] 쿠키 문자열 반환
GetCookieString(cookieMap) {
    str := ""
    for k, v in cookieMap
        str .= (str == "" ? "" : "; ") . k . "=" . v
    return str
}

; [함수] 헤더 설정
SetHeaders(mode := "DEFAULT", referer := "", targetMap := "") {
    global http

    ; 현재 타겟에 맞는 바구니로 스위칭
    current_cookies := targetMap

    http.SetRequestHeader("User-Agent",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )

    if (referer != "")
        http.SetRequestHeader("Referer", referer)

    if (mode == "AJAX") {
        http.SetRequestHeader("X-Requested-With", "XMLHttpRequest")
        http.SetRequestHeader("Content-Type", "application/json")
    } else if (mode == "FORM") {
        http.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
    }

    cookieStr := GetCookieString(current_cookies)
    if (cookieStr != "")
        http.SetRequestHeader("Cookie", cookieStr)
}

; [함수] 로그 기록
LogDebug(msg) {
    global DEBUG_FILE
    try {
        FileAppend(FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") " " msg "`n", DEBUG_FILE, "UTF-8")
    }
}

try {

    ; ==============================================================================
    ; [Step 1] 포털 로그인 (btcep)
    ; ==============================================================================
    current_cookies := btcep_cookies

    LogDebug("[Step 1] 포털 접속 시도...")
    url_main := "https://btcep.humetro.busan.kr/user/login.face?destination=%2Fportal%2F"
    http.Open("GET", url_main, false)
    SetHeaders("DEFAULT", , btcep_cookies)
    http.Send()
    UpdateCookies(http.GetAllResponseHeaders())

    ; 1차 인증
    LogDebug("[Step 1] 1차 인증 시도...")
    http.Open("POST", "https://btcep.humetro.busan.kr/user/checkUser.face", false)
    SetHeaders("AJAX", , btcep_cookies)
    http.Send('{"userId":"' . sUser . '","pwd":"' . sPw1 . '"}')
    UpdateCookies(http.GetAllResponseHeaders())

    http.Open("POST", "https://btcep.humetro.busan.kr/user/checkLogin.face", false)
    SetHeaders("AJAX", , btcep_cookies)
    http.Send('{"userId":"' . sUser . '","pwd":"' . sPw1 . '"}')
    UpdateCookies(http.GetAllResponseHeaders())

    ; 2차 인증
    if (sPw2 != "") {
        LogDebug("[Step 1] 2차 인증 시도...")
        http.Open("POST", "https://btcep.humetro.busan.kr/user/secondLoginYn.face", false)
        SetHeaders("AJAX", , btcep_cookies)
        http.Send('{"userId":"' . sUser . '","secondpwd":"' . sPw2 . '"}')
        UpdateCookies(http.GetAllResponseHeaders())
    }

    ; 최종 로그인 프로세스
    LogDebug("[Step 1] 로그인 프로세스 완료 요청...")
    finalUrl := "https://btcep.humetro.busan.kr/user/loginProcess.face?destination=%2Fportal%2F"
    postData := "username=" . sUser . "&userId=" . sUser . "&password=" . EncodeURI(sPw1)

    http.Open("POST", finalUrl, false)
    http.Option[6] := false ; No Auto Redirect
    SetHeaders("FORM", url_main, btcep_cookies)
    http.Send(postData)
    UpdateCookies(http.GetAllResponseHeaders())

    if (http.Status != 302) {
        msgbox("로그인 실패 (Status: " http.Status ")")
        ExitApp(1)
    }

    sso_exchange_url := http.GetResponseHeader("Location")
    LogDebug("[Step 1] 로그인 성공")

    ; ==============================================================================
    ; [Step 2] SSO 환전 (btcep -> niw)
    ; ==============================================================================
    LogDebug("[Step 2] SSO 환전 시도...")

    ; 바구니 교체: niw_cookies 사용
    current_cookies := niw_cookies

    http.Open("GET", sso_exchange_url, false)
    http.Option[6] := true ; Auto Redirect ON (niw까지)

    ; 요청 시 쿠키는 btcep_cookies를 사용해야 함 (SSO 토큰 전달)
    SetHeaders("DEFAULT", "https://btcep.humetro.busan.kr/", niw_cookies)

    http.Send()
    UpdateCookies(http.GetAllResponseHeaders())
    LogDebug("[Step 2] 환전 완료.")

    /*
    ; ==============================================================================
    ; [Step 2.5] 메인포털 후속 절차
    ; ==============================================================================
    LogDebug("[Step 2.5-1] 메인포털 후속 절차 시도1...")
    
    ; 바구니 교체: btcep_cookies 사용
    current_cookies := btcep_cookies
    
    http.Open("GET", "https://btcep.humetro.busan.kr/", false)
    http.Option[6] := false
    SetHeaders("DEFAULT", "https://btcep.humetro.busan.kr/", btcep_cookies)
    
    http.Send()
    if (http.Status != 302) {
        msgbox("뭔가 실패1 (Status: " http.Status ")")
        ExitApp(1)
    }
    
    LogDebug("[Step 2.5-2] 메인포털 후속 절차 시도2...")
    
    http.Open("GET", "https://btcep.humetro.busan.kr/portal", false)
    SetHeaders("DEFAULT", "https://btcep.humetro.busan.kr/", btcep_cookies)
    
    http.Send()
    if (http.Status != 200) {
        msgbox("뭔가 실패2 (Status: " http.Status ")")
        ExitApp(1)
    }
    
    current_cookies := niw_cookies
    http.Option[6] := true
    */
    ; ==============================================================================
    ; [Step 3] ERP 포털 접속을 위한 준비
    ; ==============================================================================
    LogDebug("[Step 3] ERP 포털 K값 준비...")
    timestamp := A_TickCount
    niwUrl := "https://niw.humetro.busan.kr/sso/index.jsp?callType=callMenu&menuType=fwdlogin&bw=chrome&_=" timestamp

    http.Open("GET", niwUrl, false)
    SetHeaders("DEFAULT", "https://btcep.humetro.busan.kr/", niw_cookies)
    http.Send()
    UpdateCookies(http.GetAllResponseHeaders())

    ; ==============================================================================
    ; [Step 4] erpep.jsp 호출 -> K 추출
    ; ==============================================================================

    LogDebug("[Step 4] K값 추출 시도...")
    erpep_url := "https://niw.humetro.busan.kr/erpep.jsp"

    http.Open("GET", erpep_url, false)
    http.Option[6] := false ; Auto Redirect OFF
    SetHeaders("DEFAULT", "https://btcep.humetro.busan.kr/portal/default/main/erpportal.page", niw_cookies)
    http.Send()
    UpdateCookies(http.GetAllResponseHeaders())

    html := http.ResponseText
    k_val := ""
    if RegExMatch(html, "let authCd = '([^']+)';", &match) {
        k_val := match[1]
        LogDebug("[Step 4] K값 획득 성공: " k_val)
    } else {
        msgbox("K값 추출 실패")
        ExitApp(1)
    }

    ; ==============================================================================
    ; [Step 5] SSOData 호출 (niw -> ep)
    ; ==============================================================================

    LogDebug("[Step 5] SSOData 호출 (ep 세션 생성)...")

    ; 바구니 교체: ep_cookies 사용
    current_cookies := ep_cookies

    Sleep 500
    sso_data_url :=
        "http://ep.humetro.busan.kr/irj/servlet/prt/portal/prtroot/kr.busan.humetro.cbo.ep.ssoLogin.SSOData"
    sso_payload := '{"I_SABUN":"' k_val '","I_RET":""}'

    MaxRetries := 3
    dynamic_pw := ""

    ; [핵심 포인트 2] 실패 시 넉넉한 간격을 두고 재시도하는 루프 도입
    loop MaxRetries {
        http.Open("POST", sso_data_url, false)
        SetHeaders(, kURL := "http://ep.humetro.busan.kr/irj/portal?K=" k_val, ep_cookies)
        http.SetRequestHeader("Origin", "http://ep.humetro.busan.kr")
        http.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded; charset=UTF-8")

        http.Send(sso_payload)
        UpdateCookies(http.GetAllResponseHeaders())

        sso_resp := http.ResponseText
        sso_parts := StrSplit(sso_resp, "@")

        ; 정상적으로 @ 기준 4개 이상의 파츠가 파싱되었다면 성공
        if (sso_parts.Length >= 4) {
            dynamic_pw := Trim(sso_parts[4], "`r`n ")
            LogDebug("[Step 5] 동적 비밀번호 획득 성공 (" A_Index "회차): " dynamic_pw)
            break
        } else {
            LogDebug("[Step 5] " A_Index "회차 시도 실패. 1초 대기 후 재시도...")
            Sleep(1000) ; 실패했다면 1초 푹 쉬고 다시 찌릅니다.
        }
    }

    ; 3번 다 실패했을 경우의 에러 처리
    if (dynamic_pw == "") {
        errMsg := "SSOData 응답 형식이 예상과 다릅니다 (동적 암호 추출 최종 실패)."
        LogDebug("오류: " errMsg "`n최종 응답: " sso_resp)
        FileAppend('{"error": "' errMsg '"}', "*")
        ExitApp(1)
    }

    ; 확인
    sap_cookie_key := "com.sap.engine.security.authentication.original_application_url"
    if !ep_cookies.Has(sap_cookie_key) || !InStr(ep_cookies[sap_cookie_key], "POST#") {
        LogDebug("경고: POST 모드 전환 확인 불가 (계속 진행)")
    }

    ; ==============================================================================
    ; [Step 6] MYSAPSSO2 생성 (ep)
    ; ==============================================================================
    LogDebug("[Step 6] 최종 로그인 (MYSAPSSO2 획득)...")

    ep_login_url := "http://ep.humetro.busan.kr/irj/portal?K=" k_val

    ; dynamic_pw
    login_post := "login_submit=on&login_do_redirect=1&no_cert_storing=on&j_user=" sUser "&j_password=" EncodeURI(
        dynamic_pw) "&nocomp=&newPw=&pwCheck=&WorkNo=&Name=&phoneNum="

    http.Open("POST", ep_login_url, false)
    http.Option[6] := false
    SetHeaders("FORM", ep_login_url, ep_cookies)
    http.SetRequestHeader("Origin", "http://ep.humetro.busan.kr")

    http.Send(login_post)
    UpdateCookies(http.GetAllResponseHeaders())

    if !ep_cookies.Has("MYSAPSSO2") {
        LogDebug("오류: MYSAPSSO2 획득 실패")
        msgbox("MYSAPSSO2 획득 실패")
        ExitApp(1)
    }

    LogDebug("성공: MYSAPSSO2 획득 완료")

    ; ==============================================================================
    ; [Step 7] 사용자 정보(innerpage) 추출 및 동적 쿠키 조립
    ; ==============================================================================
    LogDebug("[Step 7] 사용자 정보 추출 시도...")

    WID := DateDiff(A_NowUTC, "19700101000000", "Seconds") * 1000
    innerUrl :=
        "http://ep.humetro.busan.kr/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fhumetro!2fdesktop!2fdesktop.default!2fframeworkPages!2flayout.framework!2fcom.sap.portal.innerpage?K=" k_val "&windowId=WID" WID

    http.Open("GET", innerUrl, false)
    SetHeaders("DEFAULT", "http://ep.humetro.busan.kr/irj/portal", ep_cookies)
    http.Send()
    html_inner := http.ResponseText

    userInfoStr := ""
    if RegExMatch(html_inner, 'var\s+userinfo\s*=\s*"UserMasterEntity:\{(.*?)\}"', &match) {
        userInfoStr := match[1]
        LogDebug("[Step 7] 추출 성공: " userInfoStr)
    } else {
        MsgBox("HTML에서 사용자 정보를 찾을 수 없습니다.")
        ExitApp(1)
    }

    JSEscape(str) {
        static doc := ComObject("htmlfile")
        doc.write('<meta http-equiv="X-UA-Compatible" content="IE=edge">')
        return doc.parentWindow.escape(str)
    }

    cookieData :=
        "urn%253Akr.busan.humetro%253Anavigation%2526title%3DERP%25uD3EC%25uD138%25uC2DC%25uC2A4%25uD15C%3B%20"
    cookieData .=
        "urn%253Akr.busan.humetro%253Anavigation%2526pcdid%3DROLES%253A//portal_content/humetro/role/home/role.01/workset.home/page.06"
    allowedKeys :=
        ",SABUN,NAME,BUSEO_CODE,BUSEO_NAME,JIKGUB_CODE,JIKGUB_NAME,JIKWI_CODE,JIKWI_NAME,JIKYEL_CODE,JIKYEL_NAME,"

    userParts := StrSplit(userInfoStr, ",")
    for _, pair in userParts {
        if (pair == "")
            continue
        kv := StrSplit(pair, "=")
        if (kv.Length < 2)
            continue

        k := kv[1], v := kv[2]
        if InStr(allowedKeys, "," k ",") {
            encKey := JSEscape("urn:kr.busan.humetro:userinfo&" . k)
            encVal := JSEscape(v)
            cookieData .= "%3B%20" . EncodeURI(encKey) . "%3D" . EncodeURI(encVal)
        }
    }

    ep_cookies["SAPPORTALSDB0"] := cookieData
    LogDebug("[Step 7] 동적 SAPPORTALSDB0 조립 및 주입 완료")

    ; ==============================================================================
    ; [Step 8] SearchValueData 세션 예열 (SESS_ 변수 초기화)
    ; ==============================================================================
    LogDebug("[Step 8] SearchValueData 세션 예열 시작...")

    ; 1. [Step 7]에서 추출한 userInfoStr에서 필요한 값 뽑아내기
    val_sabun := ""
    val_name := ""
    val_buseo_cd := ""
    val_buseo_nm := ""

    if RegExMatch(userInfoStr, "SABUN=([^,]+)", &m)
        val_sabun := m[1]
    if RegExMatch(userInfoStr, "NAME=([^,]+)", &m)
        val_name := m[1]
    if RegExMatch(userInfoStr, "BUSEO_CODE=([^,]+)", &m)
        val_buseo_cd := m[1]
    if RegExMatch(userInfoStr, "BUSEO_NAME=([^,]+)", &m)
        val_buseo_nm := m[1]

    LogDebug("예열 파라미터: 사번=" val_sabun ", 이름=" val_name ", 부서=" val_buseo_cd)

    ; 2. 페이로드(Body) JSON 조립
    warmup_payload := '{"AJAX_TYPE":"HEAD","SABUN":"' val_sabun '","NAME":"' val_name '","BUSEO_CODE":"' val_buseo_cd '","BUSEO_NAME":"' val_buseo_nm '"}'

    ; 3. WinHttp 요청 설정
    warmup_url :=
        "http://ep.humetro.busan.kr/irj/servlet/prt/portal/prtroot/kr.busan.humetro.cbo.ep.search_value.SearchValueData"

    http.Open("POST", warmup_url, false)

    ; Referer 주소에 [Step 4]에서 구한 k_val 포함
    ref_url := "http://ep.humetro.busan.kr/irj/portal?K=" . k_val
    SetHeaders("DEFAULT", ref_url, ep_cookies)

    ; (기존 SetHeaders 함수의 AJAX 모드가 application/json이므로, 이 부분만 명시적으로 덮어씌웁니다)
    http.SetRequestHeader("Accept", "*/*")
    http.SetRequestHeader("Cache-Control", "no-cache")
    http.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded; charset=UTF-8")
    http.SetRequestHeader("X-Requested-With", "XMLHttpRequest")

    ; 4. 요청 발사 및 쿠키 갱신
    http.Send(warmup_payload)
    UpdateCookies(http.GetAllResponseHeaders())

    LogDebug("[Step 8] SearchValueData 예열 완료! 상태코드: " http.Status)

    ; ==============================================================================
    ; [Step 9] CDP를 이용한 크롬/엣지 브라우저 실행 및 쿠키 주입 (디버깅 모드 전용)
    ; ==============================================================================
    if (!A_IsCompiled) {
        LogDebug("[Step 9] [디버그 모드] 단독 실행용 CDP 브라우저 실행 및 쿠키 주입 시작...")

        edgePath := "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
        profile := A_Temp "\edge_cookie_profile" A_TickCount

        try {
            ; 1. 빈 페이지로 브라우저 실행
            runArgs := ' --remote-debugging-port=9222 --user-data-dir="' profile '"'
                . ' --no-first-run --no-default-browser-check --disable-default-apps'
                . ' about:blank'

            Run(Format('"{1}" {2}', edgePath, runArgs), , "max")
            Browser := Chrome([], , , 9222)
            Page := Browser.GetPage()
            Page.Call("Network.enable")

            ; 2. ep_cookies 바구니에 있는 모든 쿠키를 브라우저에 강제 주입
            for cookieName, cookieValue in ep_cookies {
                cookieParams := Map(
                    "name", cookieName,
                    "value", cookieValue,
                    "path", "/"
                )

                ; MYSAPSSO2와 우리가 만든 SAPPORTALSDB0는 서브도메인을 아우르도록 설정
                if (cookieName = "MYSAPSSO2" || cookieName = "SAPPORTALSDB0") {
                    cookieParams["domain"] := ".humetro.busan.kr"
                } else {
                    cookieParams["domain"] := "ep.humetro.busan.kr"
                }

                Page.Call("Network.setCookie", cookieParams)
            }
            LogDebug("[Step 9] [디버그 모드] CDP 쿠키 주입 완료.")

            DirectURL :=
                "http://ep.humetro.busan.kr/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fhumetro!2frole!2fmaintenance!2frole.09!2fworkset.07!2fworkset.01!2fworkset.03!2fiview.02?sapDocumentRenderingMode=EmulateIE8"
            Page.Call("Page.navigate", Map("url", DirectURL))

            LogDebug("[Step 9] [디버그 모드] 업무일지 리스트 페이지 호출 성공. 화면을 확인하세요.")

            LogDebug("========== [디버그 모드] 세션BG 단독 테스트 완료 ==========")
            ExitApp ; 디버그 단독 실행 시 브라우저만 띄우고 종료

        } catch as err {
            LogDebug("CDP 브라우저 제어 중 오류 발생: " err.Message)
            MsgBox("브라우저 실행 중 오류가 발생했습니다.`n" err.Message, "CDP Error")
            ExitApp(1)
        }
    } else {
        LogDebug("[Step 9] [운영 모드] 컴파일 실행이므로 브라우저 띄우기 생략 (세션 생성 완료)")
    }

    ; ==============================================================================
    ; [Result] 운영 모드 (컴파일 상태) 쿠키 JSON 배출 (Main.ahk에서 수신 목적)
    ; ==============================================================================
    LogDebug("[Result] 메인 프로세스로 쿠키 전달 (JSON)...")

    result := Map(
        "btcep", GetCookieString(btcep_cookies),
        "niw", GetCookieString(niw_cookies),
        "ep", GetCookieString(ep_cookies)
    )

    jsonOutput := JSON.stringify(result, 0)
    LogDebug("JSON 반환 성공. 길이: " StrLen(jsonOutput))

    ; 로그 출력 (UTF-8 명시)
    FileAppend(jsonOutput, "*", "UTF-8")
    LogDebug("========== [운영 모드] 세션BG.exe 처리 완료. 메인으로 반환 ==========")

    ; 메인 스크립트가 파이프 데이터를 수신했다는 stdin 확인 시 까지 대기
    ForceExit() {
        LogDebug("[Result] stdin 대기 타임아웃(3초) 발생. 프로세스 강제 종료.")
        ExitApp(1)
    }
    SetTimer(ForceExit, 3000) ; 3초 타임아웃 설정

    try {
        stdin := FileOpen("*", "r", "UTF-8")
        loop {
            line := Trim(stdin.ReadLine())
            if (line == "COOKIE_SAVED") {
                LogDebug("[Result] 메인 프로세스(Main.ahk)에서 쿠키 저장 완료 확인. 프로세스 종료.")
                break
            }
        }
    } catch as err {
        LogDebug("[Result] stdin 대기 중 오류 발생: " err.Message)
    }

    SetTimer(ForceExit, 0) ; 성공 시 타이머 해제
    ExitApp

} catch as e {
    MsgBox A_Clipboard := e.Message
    LogDebug("치명적 오류`n" e.Message "`n")
    FileAppend('{"error": "' . e.Message . '"}', "*")
    ExitApp(1)
}

DumpCookies(mapObj, title) {
    out := "=== " title " ===`n"
    if (mapObj.Count = 0)
        return out . "(없음)`n"

    for k, v in mapObj
        out .= k . "=" . v . "`n"

    return out
}

EncodeURI(str) {
    static doc := ComObject("htmlfile")
    doc.write('<meta http-equiv="X-UA-Compatible" content="IE=edge">')
    return doc.parentWindow.encodeURIComponent(str)
}
