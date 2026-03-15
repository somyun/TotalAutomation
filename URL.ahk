#Requires AutoHotkey v2.0

; URL 및 스프레드시트 상수
global WebAppURL :=
    "https://script.google.com/macros/s/AKfycbyoSMf94VffKSvIoBNJHKkQqY213h6M9KhTSBJ1BK9ed8dW64d50ZbjGWGu4n31bJB-/exec"
global TARGET_SPREADSHEET_ID := "19rgzRnTQtOwwW7Ts5NbBuItNey94dAZsEnO7Tk0cm6s"

; URL 인코딩 함수 (UTF-8)
URLEncode(str) {
    if (str == "")
        return ""

    byteCount := StrPut(str, "UTF-8")
    buf := Buffer(byteCount)
    StrPut(str, buf, "UTF-8")

    result := ""

    loop byteCount - 1 {
        b := NumGet(buf, A_Index - 1, "UChar")
        if (
            (b >= 0x30 && b <= 0x39) ; 0-9
            || (b >= 0x41 && b <= 0x5A) ; A-Z
            || (b >= 0x61 && b <= 0x7A) ; a-z
            || b == 0x2D || b == 0x5F || b == 0x2E || b == 0x7E
        ) {
            result .= Chr(b)
        } else {
            result .= "%" Format("{:02X}", b)
        }
    }

    return result
}
