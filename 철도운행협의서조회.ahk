#Requires AutoHotkey v2.0

class TrackAgreementService {
    static QueryUrl := "https://mis.humetro.busan.kr/FS/procSelect.do"

    static FindMatches(participantNames, targetDate := "") {
        if (targetDate = "")
            targetDate := FormatTime(DateAdd(A_Now, 1, "Days"), "yyyyMMdd")
        if !RegExMatch(targetDate, "^\d{8}$")
            throw Error("협의서 조회 날짜 형식이 올바르지 않습니다.")

        expectedNames := this.NormalizeNameSet(participantNames)
        if (expectedNames.Count = 0)
            throw Error("작업관계자 이름이 입력되지 않았습니다.")

        listResponse := SessionManager.Request(
            "POST",
            this.QueryUrl,
            this.BuildListQueryXml(targetDate),
            Map("Content-Type", "text/xml;charset=UTF-8")
        )
        listDoc := XPlatformProtocol.EnsureSuccess(listResponse.Text)
        rows := listDoc.SelectNodes(
            "//*[local-name()='Dataset' and @id='ds_out']/*[local-name()='Rows']/*[local-name()='Row']")

        matches := []
        for row in rows {
            item := this.ReadRow(row)
            agreementNo := this.Get(item, "MASTER_SEQ")
            if (agreementNo = "" || !this.ContainsDate(item, targetDate))
                continue

            detailResponse := SessionManager.Request(
                "POST",
                this.QueryUrl,
                this.BuildDetailQueryXml(agreementNo),
                Map("Content-Type", "text/xml;charset=UTF-8")
            )
            detailDoc := XPlatformProtocol.EnsureSuccess(detailResponse.Text)
            actual := this.ReadParticipants(detailDoc)
            ; 협의서 관계자 전원이 현재 UI에 입력된 이름들 안에 포함되면 일치이다.
            ; UI의 네 이름 중 일부가 협의서에 없더라도 허용한다.
            if !this.AllNamesAllowed(expectedNames, actual.Set)
                continue

            matches.Push(Map(
                "AgreementNo", agreementNo,
                "WorkName", this.Get(item, "WORK_NAME"),
                "WorkPeriod", this.Get(item, "WORK_DATE_TIME"),
                "WorkSection", this.FormatSection(item),
                "Participants", actual.Display,
                "EmployeeCount", this.Get(item, "EMP_CNT")
            ))
        }

        LogDebug("철도운행협의서 조회 완료: " targetDate ", 작업관계자 일치 " matches.Length "건")
        return matches
    }

    static BuildListQueryXml(targetDate) {
        params := Map(
            "queryId", "LA10800.searchList",
            "IN_START_DATE", targetDate,
            "IN_END_DATE", targetDate
        )
        return XPlatformProtocol.Root(XPlatformProtocol.BuildParameters(params)
            . XPlatformProtocol.BuildTransInfo("searchList", "procSelect.do", "", "", false))
    }

    static BuildDetailQueryXml(agreementNo) {
        params := Map("queryId", "LA10800.searchMaster", "MASTER_SEQ", agreementNo)
        return XPlatformProtocol.Root(XPlatformProtocol.BuildParameters(params)
            . XPlatformProtocol.BuildTransInfo("searchMaster", "procSelect.do", "", "", false))
    }

    static ReadParticipants(doc) {
        rows := doc.SelectNodes(
            "//*[local-name()='Dataset' and @id='ds_out3']/*[local-name()='Rows']/*[local-name()='Row']")
        names := []
        set := Map()
        for row in rows {
            values := this.ReadRow(row)
            displayName := Trim(this.Get(values, "EMPL_NAME"))
            normalized := this.NormalizeName(displayName)
            if (normalized = "" || set.Has(normalized))
                continue
            set[normalized] := true
            names.Push(displayName)
        }
        return {Set: set, Display: this.Join(names, ", ")}
    }

    static NormalizeNameSet(names) {
        result := Map()
        if !IsObject(names)
            return result
        for name in names {
            normalized := this.NormalizeName(name)
            if (normalized != "")
                result[normalized] := true
        }
        return result
    }

    static AllNamesAllowed(allowed, actual) {
        if (actual.Count = 0)
            return false
        for name in actual {
            if !allowed.Has(name)
                return false
        }
        return true
    }

    static ContainsDate(values, targetDate) {
        ; LA10800 실제 응답은 WORK_DATE1=시작일, WORK_DATE=종료일을 제공한다.
        startDate := this.FirstValue(values,
            ["WORK_DATE1", "WORK_START_DATE", "START_DATE", "WORK_STRT_DATE", "WORK_BGN_DATE"])
        endDate := this.FirstValue(values,
            ["WORK_DATE", "WORK_END_DATE", "END_DATE", "WORK_FNSH_DATE"])
        startDate := this.FirstDate(startDate)
        endDate := this.FirstDate(endDate)
        if (startDate != "" || endDate != "") {
            if (startDate = "")
                startDate := endDate
            if (endDate = "")
                endDate := startDate
            return startDate <= targetDate && targetDate <= endDate
        }

        period := this.Get(values, "WORK_DATE_TIME")
        dates := []
        position := 1
        while RegExMatch(period, "(\d{4})\D+(\d{1,2})\D+(\d{1,2})", &match, position) {
            dates.Push(match[1] Format("{:02}", Number(match[2])) Format("{:02}", Number(match[3])))
            position := match.Pos + match.Len
        }
        if (dates.Length = 0)
            return false
        return dates[1] <= targetDate && targetDate <= dates[dates.Length]
    }

    static FirstDate(value) {
        if RegExMatch(value, "(\d{4})\D*(\d{1,2})\D*(\d{1,2})", &match)
            return match[1] Format("{:02}", Number(match[2])) Format("{:02}", Number(match[3]))
        return ""
    }

    static FirstValue(values, keys) {
        for key in keys {
            value := this.Get(values, key)
            if (value != "")
                return value
        }
        return ""
    }

    static FormatSection(values) {
        startName := this.Get(values, "START_STNM")
        endName := this.Get(values, "END_STNM")
        if (startName = "")
            return endName
        if (endName = "" || endName = startName)
            return startName
        return startName " ~ " endName
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

    static NormalizeName(value) {
        return StrLower(RegExReplace(Trim(value), "\s+", ""))
    }

    static Join(values, delimiter) {
        output := ""
        for value in values
            output .= (output = "" ? "" : delimiter) value
        return output
    }
}
