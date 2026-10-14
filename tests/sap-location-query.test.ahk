#Requires AutoHotkey v2.0 32-bit
#SingleInstance Off
#NoTrayIcon
#Include %A_ScriptDir%\..\Lib\JSON.ahk
#Include %A_ScriptDir%\..\Lib\SapRfcClient.ahk
#Include %A_ScriptDir%\..\Lib\SapLocationOrders.ahk

class FakeLocationRfc {
    __New() {
        this.calls := [], this.badPage := false, this.returnError := false
        this.rows := []
        Loop 3 {
            row := Map()
            for field in ["ORDERID", "ORDER_TYPE", "MN_WK_CTR", "PLANPLANT", "START_DATE", "ENTER_DATE", "U_STATUS", "S_STATUS", "SHORT_TEXT", "FUNCLOC", "EQUIPMENT", "MAINTPLANT", "MAINTLOC", "MAINTROOM", "PLSECTN"]
                row[field] := ""
            row["ORDERID"] := Format("{:012}", A_Index)
            row["ORDER_TYPE"] := "PM06", row["MN_WK_CTR"] := "5129", row["PLANPLANT"] := "5010"
            row["START_DATE"] := "20260101", row["U_STATUS"] := A_Index = 2 ? "완료" : "작업"
            row["SHORT_TEXT"] := "원본 작업_장소", row["FUNCLOC"] := "시설"
            this.rows.Push(row)
        }
    }
    Function(name) {
        if name != "BAPI_ALM_ORDERHEAD_GET_LIST"
            throw Error("Unexpected function")
        this.ranges := {Rows: []}, this.display := Map(), this.nav := Map()
        this.result := {Rows: [], Index: 1}
        return 1
    }
    Table(handle, name) => name = "IT_RANGES" ? this.ranges : this.result
    Append(table) {
        row := Map(), table.Rows.Push(row)
        return row
    }
    Set(handle, name, value) => handle[name] := value
    Int(handle, name, value) => handle[name] := value
    Structure(handle, name) => name = "DISPLAY_PARAMETERS" ? this.display : this.nav
    Invoke(handle) {
        page := this.display["SHOW_PAGE_NUMBER"], length := this.display["PAGELENGTH"]
        ids := Map()
        for range in this.ranges.Rows
            if range["FIELD_NAME"] = "OPTIONS_FOR_ORDERID"
                ids[range["LOW_VALUE"]] := true
        selected := []
        for row in this.rows
            if ids.Count = 0 || ids.Has(row["ORDERID"])
                selected.Push(row)
        this.result.Rows := []
        for row in selected {
            if A_Index > (page - 1) * length && A_Index <= page * length
                this.result.Rows.Push(row)
        }
        this.nav := Map("NUMBER_OF_HITS", selected.Length, "ACTUAL_PAGE", this.badPage && page > 1 ? 1 : page,
            "LAST_PAGE", Ceil(selected.Length / length))
        this.calls.Push(this.ranges.Rows)
    }
    Returns(handle) {
        if this.returnError
            throw Error("BAPI error")
    }
    Count(table) => table.Rows.Length
    Move(table, index) => table.Index := index + 1
    Get(handle, name) => handle is Map ? String(handle[name]) : String(handle.Rows[handle.Index][name])
}

Assert(condition, label) {
    if !condition
        throw Error(label)
}
HasRange(ranges, field, low) {
    for range in ranges
        if range["FIELD_NAME"] = field && range["LOW_VALUE"] = low
            return true
    return false
}
exitCode := 0
try {
    rfc := FakeLocationRfc()
    rows := QueryLocationOrders(rfc, "5129", "2026", 1)
    Assert(rows.Length = 2, "Active rows/paging")
    Assert(HasRange(rfc.calls[1], "OPTIONS_FOR_STATUS_INCLUSIVE", "작업"), "Active status range")
    rows := QueryExistingLocationOrders(rfc, "5129", ["1", "2", "3"])
    Assert(rows.Length = 3 && rows[2]["U_STATUS"] = "완료", "Completed order retained")
    Assert(!HasRange(rfc.ranges.Rows, "OPTIONS_FOR_STATUS_INCLUSIVE", "작업"), "Existing lookup without active filter")
    Assert(HasRange(rfc.ranges.Rows, "OPTIONS_FOR_ORDERID", "000000000001"), "Padded exact order ID")
    Assert(rows[1]["SHORT_TEXT"] = "원본 작업_장소" && rows[1]["FUNCLOC"] = "시설", "Original fields preserved")
    rows := QuerySapLocationOrders(rfc, "5129", "19000101", "99991231", ["000000000001", "000000000002", "000000000003"], false, 1)
    Assert(rows.Length = 3, "All-state paging")
    rfc.badPage := true, failed := false
    try QueryLocationOrders(rfc, "5129", "2026", 1)
    catch
        failed := true
    Assert(failed, "Page omission must fail")
    rfc.badPage := false, rfc.returnError := true, failed := false
    try QueryExistingLocationOrders(rfc, "5129", ["1"])
    catch
        failed := true
    Assert(failed, "BAPI errors must fail")
    rfc.returnError := false
    Assert(QueryExistingLocationOrders(rfc, "5129", []).Length = 0, "Empty ID list")
    ids := []
    Loop 101
        ids.Push(String(A_Index))
    before := rfc.calls.Length
    QueryExistingLocationOrders(rfc, "5129", ids)
    Assert(rfc.calls.Length - before = 2, "Batch exact queries")
    FileAppend("PASS RFC ranges/completed orders/raw identity/paging/error handling/batching`n", "*", "UTF-8")
} catch as err {
    FileAppend("FAIL " err.Message " line=" err.Line "`n", "*", "UTF-8")
    exitCode := 1
}
ExitApp(exitCode)
