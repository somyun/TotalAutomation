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

    ; Win + Z : 자동 로그인
    static AutoLoginAction(*) {
        user := ConfigManager.CurrentUser
        if (!user || user["id"] == "") {
            MsgBox "로그인된 사용자가 없습니다."
            return
        }

        ShortcutActions.AutoERPPortal(user["id"], user["webPW"], user["pw2"])

        ; 로그인 후 포커스 문제 해결
        if !WinActive("ERP포털시스템 - 부산교통공사") {
            try {
                cUIA := UIA_Browser("ERP포털시스템 - 부산교통공사")
                cUIA.WaitElement({ ClassName: "lastestip" }, 5000)
                cUIA.send("{esc}")
            }
        }
    }

    ; Win + Alt + Z : 자동 로그인 + 일지 열기
    static AutoLoginOpenLogAction(*) {
        KeyWait "LWin"
        KeyWait "Alt"

        user := ConfigManager.CurrentUser
        if (!user || user["id"] == "") {
            MsgBox "로그인된 사용자가 없습니다."
            return
        }

        ; 브라우저 찾기
        browserExe := ""
        for exe in ["chrome.exe", "msedge.exe"] {
            if ProcessExist(exe) {
                browserExe := exe
                break
            }
        }

        ; 로그인 실행
        cUIA := ShortcutActions.AutoERPPortal(user["id"], user["webPW"], user["pw2"], browserExe)

        ;팝업 처리
        cUIA.WaitElement({ ClassName: "lastestip" }, 5000)
        cUIA.send("{esc}")

        ;ERP포털 이동
        cUIA.WaitElement({ Type: "Link", Name: "ERP" }, 3000).Invoke()

        ; 업무일지 메뉴 이동
        cUIA.WaitElement({ Name: "업무일지 업무일지" }, 10000).Invoke()

        ; 조회 및 클릭
        targetDate := FormatTime(DateAdd(A_Now, -9, "Hours"), "yyyyMMdd")
        targetName := targetDate " " user["department"] " 업무일지" ; 부서명 동적으로

        try
            dept := user["department"]
        catch
            MsgBox "업무일지 리스트 선택 실패, 분소명이 설정되지 않은 상태입니다"

        btnName := targetDate " " dept " 업무일지"

        cUIA.WaitElement({ Name: btnName }, 5000).Click("Left")

        cUIA.FindElement({ LocalizedType: "링크", Name: "변경/조회" }).Invoke()

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

    ; ERP 로그인 로직
    static AutoERPPortal(id, pw, certi, browserExe := "msedge.exe") {
        url := " https://btcep.humetro.busan.kr/portal/default/main/erpportal.page"

        ; 브라우저 실행 (이미 열려있으면 탭 추가/활성화는 UIA_Browser가 처리하거나 사용자가 함)
        if (browserExe != "") {
            Run browserExe " " url
        } else {
            Run url
        }
        Sleep 100

        loop 100 {
            try {
                cUIA := UIA_Browser(":: 부산교통공사 포털시스템 ::")
                cUIA.WaitElement({ AutomationId: "userId" }, 2000).value := id
                cUIA.WaitElement({ AutomationId: "password" }, 1000).value := pw
                cUIA.WaitElement({ ClassName: "btn_login" }, 1000).invoke()

                ; 2차 인증
                if (certi != "") {
                    cUIA.WaitElement({ AutomationId: "certi_num" }, 2000).value := certi
                    cUIA.WaitElement({ ClassName: "btn_blue" }, 1000).invoke()
                }

                cUIA.WaitTitleChange(":: 부산교통공사 ::", 5000)
                return cUIA
            }
            catch {
                try {
                    cUIA := UIA_Browser("ERP포털시스템 - 부산교통공사")
                    return cUIA
                }
            }
            Sleep 100
        }
        return ""
    }
}
