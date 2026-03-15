#Requires AutoHotkey v2.0

; ==============================================================================
; 엑셀 자동화 핸들러 클래스
; 전역 변수 대신 정적 프로퍼티를 사용하여 엑셀 상태를 관리합니다.
; ==============================================================================
class ExcelHandler {
    static App := ""
    static Workbook := ""
    static Sheet := ""

    ; 엑셀 연결 (Win+Alt+C)
    static ConnectToActiveSheet() {
        if !WinActive("ahk_class XLMAIN") {
            MsgBox "엑셀 창이 아닙니다`n일지내용이 적힌 엑셀창에서 단축키를 실행하세요", "복사 실패", "Iconx"
            return
        }

        try {
            MainHwnd := WinExist("A")
            try {
                SheetHwnd := ControlGetHwnd("EXCEL71", MainHwnd)
            } catch {
                MsgBox "엑셀 시트 영역(EXCEL71)을 찾을 수 없습니다.`n엑셀 파일이 열려있는지 확인해주세요.", "오류", "IconX"
                return
            }

            IID_IDispatch := "{00020400-0000-0000-C000-000000000046}"
            GUID := Buffer(16)
            DllCall("Ole32\CLSIDFromString", "Str", IID_IDispatch, "Ptr", GUID)

            pacc := 0
            if DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", SheetHwnd, "UInt", -16, "Ptr", GUID, "Ptr*", &pacc) !=
            0 {
                MsgBox "엑셀 객체 연결 실패.`n관리자 권한 문제이거나 셀 편집 중일 수 있습니다.", "오류", "IconX"
                return
            }

            WindowObject := ComValue(9, pacc, 1)

            ; 정적 프로퍼티에 저장
            ExcelHandler.App := WindowObject.Application
            ExcelHandler.Workbook := ExcelHandler.App.ActiveWorkbook
            ExcelHandler.Sheet := ExcelHandler.App.ActiveSheet

            MsgBox "복사 대기 완료`n`n파일: " ExcelHandler.Workbook.Name "`n시트: " ExcelHandler.Sheet.Name, "복사 성공", "Iconi"

        } catch as e {
            MsgBox "연결 중 오류가 발생했습니다.`n" e.Message, "오류", "Iconx"
        }
    }

    ; 웹 테이블 -> 엑셀 변환 (Win+Alt+A)
    static WebTableToExcel() {
        try {
            ; 현재 활성창 기준 UIA 스캔
            targetWin := WinExist("A")
            api := UIA_Browser(targetWin)
            table := api.FindElement({ AutomationId: "tb10" })
        } catch {
            MsgBox "업무일지 일반업무 창이 아니거나 일지 로딩에 실패하였습니다", "일반업무도우미", "iconx"
            return
        }

        try {
            xl := ComObject("Excel.Application")
            wb := xl.Workbooks.Add()
            sh := wb.Sheets(1)
            sh.name := "일반업무"
            sh.Cells.Font.Size := 10

            ; 점검구분 유효성검사
            try
                sh.Range("B2:B100").Validation.Add(3, 1, 1, "전체,내부업무,점검업무,유지보수,협조사항")
            catch as e
                OutputDebug "유효성 검사 설정 실패: " e.Message

            xl.Visible := True
            xl.WindowState := -4143 ; Normal

            ; 화면 우측 하단으로 이동 (해상도에 따라 다를 수 있음, 일단 원본 로직 유지)
            try WinMove(, , 1300, 700, "ahk_id " xl.Hwnd)

            sh.Range("G:G, I:I").NumberFormat := '0000"-"00"-"00'
            sh.Range("H:H, J:J").NumberFormat := '00":"00'

            xl.ScreenUpdating := False

            for Row in table.FindAll({ LocalizedType: "행" }) {
                RowIndex := A_Index
                for Cell in Row.FindAll({ Type: "DataItem" }) {
                    CellValue := Cell.Name

                    ; 날짜/시간 필드 포맷팅 (G,H,I,J 열)
                    if (A_Index >= 7 && A_Index <= 10) {
                        CellValue := StrReplace(CellValue, "-", "")
                        CellValue := StrReplace(CellValue, ":", "")
                        if StrLen(CellValue) == 6 and (A_Index == 8 or A_Index == 10)
                            CellValue := SubStr(CellValue, 1, 4)
                        if (CellValue == "")
                            CellValue := ""
                    }
                    sh.Cells(RowIndex, A_Index).Value := CellValue
                }
            }

            sh.UsedRange.EntireColumn.AutoFit()
            xl.ScreenUpdating := True

            MsgBox "일반업무 내용을 엑셀에 복사했습니다.", "완료", "Iconi"
            WinActivate("ahk_id " xl.Hwnd)

        } catch as e {
            MsgBox "엑셀 생성 중 오류 발생: " e.Message, "오류", "IconStop"
        }
    }

    ; 엑셀 -> 웹 테이블 (Win+Alt+V)
    static ExcelToWebTable() {
        if (!ExcelHandler.App || !ExcelHandler.Sheet) {
            MsgBox "엑셀이 연결되지 않았습니다. 먼저 Win+Alt+C 로 엑셀을 연결해주세요.", "오류", "Iconx"
            return
        }

        try {
            if !targetWin := WinActive("부산교통공사 - ") {
                MsgBox "업무일지 페이지가 아닙니다.", "오류", "Iconx"
                return
            }
            api := UIA_Browser(targetWin)
            table := api.FindElement({ AutomationId: "tb10" })
        } catch {
            MsgBox "일반업무 테이블(tb10)을 찾을 수 없습니다.", "오류", "Iconx"
            return
        }

        if (MsgBox("현재 일지의 일반업무에 엑셀의 내용으로 덮어씁니다.`n`n계속 하시겠습니까?", "덮어쓰기 주의", "YesNo Icon!") == "No")
            return

        try {
            Rows := table.FindAll({ LocalizedType: "행" })
            ExcelLastRow := ExcelHandler.Sheet.UsedRange.Rows.Count

            ; 부족한 행 추가
            while Rows.Length < ExcelLastRow {
                api.FindElement({ AutomationId: "btnAdd10" }).Invoke()
                Sleep 50
                Rows := table.FindAll({ LocalizedType: "행" })
                Sleep 50
            }

            for Row in Rows {
                RowIndex := A_Index
                if (RowIndex == 1) ; 헤더 스킵
                    continue

                Cells := Row.FindAll({ Type: "DataItem" })
                for Cell in Cells {
                    ColIndex := A_Index
                    ExcelValue := ExcelHandler.Sheet.Cells(RowIndex, ColIndex).Text

                    if ColIndex == 1 {
                        if !ExcelValue
                            break
                        continue ; 순번 컬럼 스킵
                    }

                    try {
                        InputControl := Cell.FindElement({})
                        ExcelHandler.SetElementValue(InputControl, ExcelValue)
                    } catch {
                        ; Read-only or empty structure
                    }
                }
            }
            MsgBox "붙여넣기 완료", "완료", "IconI"

        } catch as e {
            MsgBox "데이터 입력 중 오류 발생: " e.Message
        }
    }

    ; UIA 요소 값 설정 헬퍼
    static SetElementValue(obj, val) {
        if obj.parent.Name == val
            return

        if obj.Type == 50003 { ; ComboBox or similar
            obj.expand()
            Sleep 50
            try obj.waitelement({ Name: val }, 3000).invoke()
            Sleep 50
            obj.collapse()
            Sleep 50
        } else {
            obj.value := val
        }
    }
}

; ==============================================================================
; 단축키 동작 클래스
; 단축키 동작을 정의합니다.
; ==============================================================================
class ShortcutActions {

    ; 공통: 브라우저 실행 및 쿠키 주입 후 이동

    ; Win + Z : 자동 로그인 (일반 모드)
    static AutoLoginAction(*) {
        user := ConfigManager.CurrentUser
        if (!user || user["id"] == "") {
            MsgBox "로그인된 사용자가 없습니다."
            return
        }

        ; 1. 브라우저 실행
        url := "https://btcep.humetro.busan.kr/portal/"
        edgePath := "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
        profile := A_Temp "\edge_cookie_profile"
        runArgs := ' --remote-debugging-port=9222 --user-data-dir="' profile '"'
    . ' --no-first-run --no-default-browser-check --disable-default-apps --new-window ' url

        if FileExist(edgePath)
            Run(Format('"{1}" {2}', edgePath, runArgs),,"max" ,&browser_pid)
        else
            Run(url)

        ; 2. 쿠키 대기 (중복 로그인 방지: 세션BG 완료 시까지 대기)
        loop 10 {
            if (HeadlessAutomation.CookieStorage.Has("btcep")) {
                cookie_btcep := HeadlessAutomation.CookieStorage["btcep"]
                cookie_niw := HeadlessAutomation.CookieStorage["niw"]
                ; JSESSIONID, K, key 쿠키 유무 확인
                if (InStr(cookie_btcep, "JSESSIONID") && InStr(cookie_niw, "K=") && InStr(cookie_niw, "key="))
                    break
            }
            if A_Index == 10 {
                MsgBox "세션bg 쿠키 확보 확인 실패"
                return
            }

            Sleep 250
        }

        ; 3. UIA 로그인 진행
        loop 10 {
            try {
                browser_hwnd := WinWait("ahk_pid " browser_pid "부산교통공사", , 3)
                Loop 10 {
                    if InStr(WinGetTitle(browser_hwnd), ":: 부산교통공사 ::") {
                        MsgBox "이미로그인"
                        return
                    }
                    else if InStr(WinGetTitle(browser_hwnd), ":: 부산교통공사 포털시스템 ::") {
                        cUIA := UIA_Browser(":: 부산교통공사 포털시스템 ::")
                    }
                    else if A_Index == 10{
                        MsgBox "타이틀 매칭 실패 " WinGetTitle("A")
                        return
                    }
                    Sleep(100)
                }

                ; 아이디 입력창 대기
                cUIA.WaitElement({ AutomationId: "userId" }, 1000).Value := user["id"]
                cUIA.WaitElement({ AutomationId: "password" }, 1000).Value := user["webPW"]
                cUIA.WaitElement({ ClassName: "btn_login" }, 1000).Invoke()

                ; 2차 비밀번호
                cUIA.WaitElement({ AutomationId: "certi_num" }, 3000).Value := user["pw2"]
                cUIA.WaitElement({ ClassName: "btn_blue" }, 1000).Invoke()

                cUIA.WaitTitleChange(":: 부산교통공사 ::", 5000)

                ; 로그인 후 포커스 해결
                cUIA.WaitElement({ ClassName: "lastestip" }, 5000)
                cUIA.send("{esc}")

                break

            } catch as e {
                Sleep 250
                if A_Index == 10 {
                    MsgBox("타임아웃, cUIA 할당 실패`n" e.Message)
                    return
                }
            }
        }
    }

    ; Win + Alt + Z : 일지 열기 (앱 모드)
    static OpenLogAction(*) {
        WebAutoLogin.EnsureReady("WorkLog_View")
    }

    ; Win + Alt + A : 웹 -> 엑셀
    static ConvertExcelAction(*) {
        ExcelHandler.WebTableToExcel()
    }

    ; Win + Alt + C : 엑셀 복사
    static CopyExcelAction(*) {
        ExcelHandler.ConnectToActiveSheet()
    }

    ; Win + Alt + V : 엑셀 -> 웹
    static PasteExcelAction(*) {
        ExcelHandler.ExcelToWebTable()
    }

    ; Win + Ctrl + Esc : 종료
    static ForceExitAction(*) {
        OnExitApp()
    }
}
