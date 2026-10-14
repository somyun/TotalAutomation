QueryLocationOrders(rfc, workCenter, year, pageLength := 1000) {
    return QuerySapLocationOrders(rfc, workCenter, year "0101", year "1231", 0, true, pageLength)
}

QueryExistingLocationOrders(rfc, workCenter, orderIds) {
    rows := [], batch := []
    for orderId in orderIds {
        if !RegExMatch(String(orderId), "^\d{1,12}$")
            throw Error("오더번호 형식이 올바르지 않습니다.")
        batch.Push(SubStr("000000000000" orderId, -12))
        if batch.Length = 100 {
            for row in QuerySapLocationOrders(rfc, workCenter, "19000101", "99991231", batch, false)
                rows.Push(row)
            batch := []
        }
    }
    if batch.Length {
        for row in QuerySapLocationOrders(rfc, workCenter, "19000101", "99991231", batch, false)
            rows.Push(row)
    }
    return rows
}

; 기존 번호 조회에는 사용자상태 제한을 걸지 않아 완료된 전년도 오더도 대조한다.
QuerySapLocationOrders(rfc, workCenter, fromDate, toDate, orderIds := 0, activeOnly := true, pageLength := 1000) {
    handle := rfc.Function("BAPI_ALM_ORDERHEAD_GET_LIST")
    ranges := rfc.Table(handle, "IT_RANGES")
    for orderType in ["PM01", "PM03", "PM06"]
        AddRange(rfc, ranges, "OPTIONS_FOR_DOC_TYPE", orderType)
    AddRange(rfc, ranges, "OPTIONS_FOR_MAIN_WORKCENTER", workCenter)
    AddRange(rfc, ranges, "OPTIONS_FOR_PLANPLANT", "5010")
    requestedIds := Map()
    if IsObject(orderIds) {
        for orderId in orderIds {
            AddRange(rfc, ranges, "OPTIONS_FOR_ORDERID", orderId)
            requestedIds[orderId] := true
        }
    } else
        AddRange(rfc, ranges, "OPTIONS_FOR_START_DATE", fromDate, "BT", toDate)
    if activeOnly
        AddRange(rfc, ranges, "OPTIONS_FOR_STATUS_INCLUSIVE", "작업")
    for flag in ["SHOW_OPEN_DOCUMENTS", "SHOW_DOCUMENTS_IN_PROCESS", "SHOW_COMPLETED_DOCUMENTS", "SHOW_HISTORICAL_DOCUMENTS"]
        AddRange(rfc, ranges, flag, "X")
    AddRange(rfc, ranges, "SHOW_DOCS_WITH_FROM_DATE", fromDate)
    AddRange(rfc, ranges, "SHOW_DOCS_WITH_TO_DATE", toDate)
    display := rfc.Structure(handle, "DISPLAY_PARAMETERS")
    rfc.Int(display, "PAGELENGTH", pageLength)
    rows := [], seen := Map(), page := 1, expectedHits := -1, fetched := 0
    loop {
        rfc.Int(display, "SHOW_PAGE_NUMBER", page)
        rfc.Invoke(handle)
        rfc.Returns(handle)
        table := rfc.Table(handle, "ET_RESULT"), count := rfc.Count(table)
        nav := rfc.Structure(handle, "NAVIGATION_DATA")
        hits := Integer(rfc.Get(nav, "NUMBER_OF_HITS"))
        actualPage := Integer(rfc.Get(nav, "ACTUAL_PAGE"))
        lastPage := Integer(rfc.Get(nav, "LAST_PAGE"))
        if expectedHits = -1
            expectedHits := hits
        if hits != expectedHits || (hits > 0 && (actualPage != page || lastPage < page || count = 0))
            throw Error("조회 도중 목록이 변경되었거나 페이지가 누락되었습니다.")
        Loop count {
            rfc.Move(table, A_Index - 1), row := Map()
            for field in ["ORDERID", "ORDER_TYPE", "MN_WK_CTR", "PLANPLANT", "START_DATE", "ENTER_DATE", "U_STATUS", "S_STATUS", "SHORT_TEXT", "FUNCLOC", "EQUIPMENT", "MAINTPLANT", "MAINTLOC", "MAINTROOM", "PLSECTN"]
                row[field] := rfc.Get(table, field)
            if row["PLANPLANT"] != "5010" || row["MN_WK_CTR"] != workCenter
                || !RegExMatch(row["ORDER_TYPE"], "^PM(01|03|06)$")
                throw Error("서버 응답이 요청한 조회 조건과 다릅니다.")
            if IsObject(orderIds) {
                if !requestedIds.Has(row["ORDERID"])
                    throw Error("요청하지 않은 오더번호가 반환되었습니다.")
            } else if row["START_DATE"] < fromDate || row["START_DATE"] > toDate
                throw Error("서버 응답의 기본시작일이 조회기간과 다릅니다.")
            ; 사용자 상태는 복수 상태 표시일 수 있으므로 '작업' 토큰을 대조한다.
            if activeOnly && !RegExMatch(row["U_STATUS"], "(^|\s)작업(\s|$)")
                continue
            orderId := row["ORDERID"]
            if !RegExMatch(orderId, "^\d+$")
                throw Error("오더번호가 없는 응답입니다.")
            if !seen.Has(orderId) {
                seen[orderId] := true
                rows.Push(row)
            }
        }
        fetched += count
        if page >= lastPage
            break
        page++
    }
    if fetched != expectedHits
        throw Error("전체 오더 목록을 받지 못했습니다.")
    return rows
}
