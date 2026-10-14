class LocationImportService {
    static Query(workCenter, employeeId) => this.RunQuery(workCenter, employeeId)["rows"]
    static Maintenance(workCenter, employeeId, orderIds) => this.RunQuery(workCenter, employeeId, orderIds)
    static RunQuery(workCenter, employeeId, orderIds := 0) {
        if !RegExMatch(workCenter, "^\d+$") || !RegExMatch(employeeId, "^\d+$")
            throw Error("로그인 사용자의 작업장 정보를 확인해주세요.")
        request := Map("command", IsObject(orderIds) ? "LocationMaintenance" : "Locations",
            "workCenter", workCenter, "employeeId", employeeId, "year", FormatTime(A_Now, "yyyy"))
        if IsObject(orderIds)
            request["orders"] := orderIds
        result := SapRfcService.Query(request)
        if IsObject(orderIds) && (!result.Has("existingRows") || !(result["existingRows"] is Array)
            || !result.Has("currentRows") || !(result["currentRows"] is Array) || !result.Has("year"))
            throw Error("SAP 오더 갱신 결과 형식이 올바르지 않습니다.")
        return result
    }
}
