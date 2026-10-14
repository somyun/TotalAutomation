class MaterialLocationService {
    static Query(employeeId, locations, searchText) {
        if !(locations is Array) || !locations.Length || !RegExMatch(employeeId, "^\d+$")
            throw Error("저장위치와 로그인 사용자를 확인해주세요.")
        return SapRfcService.Query(Map("command", "MaterialLocations", "employeeId", employeeId,
            "locations", locations, "search", searchText))["rows"]
    }
}
