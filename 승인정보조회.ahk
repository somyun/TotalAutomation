#Requires AutoHotkey v2.0

class ApprovalInfoService {
    static QueryUrl := "https://mis.humetro.busan.kr/FS/selectList.do"

    static FindFirst(driverName, startDate := "", endDate := "") {
        driverName := this.NormalizeDriverName(driverName)
        if (driverName = "")
            throw Error("운전자 이름이 없습니다.")
        if (startDate = "")
            startDate := FormatTime(A_Now, "yyyyMMdd")
        if (endDate = "")
            endDate := FormatTime(DateAdd(A_Now, 1, "Days"), "yyyyMMdd")
        if !RegExMatch(startDate, "^\d{8}$") || !RegExMatch(endDate, "^\d{8}$")
            throw Error("승인정보 조회 날짜 형식이 올바르지 않습니다.")

        response := SessionManager.Request(
            "POST",
            this.QueryUrl,
            XPlatformProtocol.BuildApprovalQueryXml(startDate, endDate),
            Map("Content-Type", "text/xml;charset=UTF-8")
        )
        doc := XPlatformProtocol.EnsureSuccess(response.Text)
        selection := this.SelectFirst(doc, driverName)
        first := selection.First

        if !IsObject(first)
            throw Error("오늘부터 내일까지 해당 운전원의 철도장비운행 승인정보를 찾지 못했습니다.")

        LogDebug("승인정보 조회 완료: 일치 " selection.Count "건, 첫 번째 결과 적용")
        return first
    }

    static SelectFirst(doc, normalizedDriverName) {
        rows := doc.SelectNodes(
            "//*[local-name()='Dataset' and @id='ds_out']/*[local-name()='Rows']/*[local-name()='Row']")

        first := 0
        matchCount := 0
        for row in rows {
            values := this.ReadRow(row)
            actualName := this.NormalizeDriverName(this.Get(values, "DRIV_NAME"))
            workName := Trim(this.Get(values, "WORK_NAME"))
            if (actualName != normalizedDriverName)
                continue
            if !InStr(workName, "철도장비운행")
                continue

            matchCount += 1
            if !IsObject(first) {
                first := Map(
                    "ApprovalNo", this.Get(values, "APPR_NUMB"),
                    "Dept", this.Get(values, "CNAP_BSNM"),
                    "Approver", this.Get(values, "CNAP_NAME")
                )
            }
        }

        return {First: first, Count: matchCount}
    }

    static ReadRow(row) {
        values := Map()
        for col in row.SelectNodes("*[local-name()='Col']")
            values[col.GetAttribute("id")] := col.Text
        return values
    }

    static Get(values, key) {
        return values.Has(key) ? Trim(values[key]) : ""
    }

    static NormalizeDriverName(value) {
        return StrLower(RegExReplace(Trim(value), "\s+", ""))
    }
}
