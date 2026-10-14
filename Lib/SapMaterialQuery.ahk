; RFC_READ_TABLE을 사용한 제한된 조회. 인증정보/SDK 메시지는 출력하지 않는다.
SapReadTable(rfc, tableName, fieldNames, options, maxRows := 20000) {
    rows := [], skip := 0, pageSize := 500
    Loop {
        handle := rfc.Function("RFC_READ_TABLE")
        rfc.Set(handle, "QUERY_TABLE", tableName)
        rfc.Int(handle, "ROWSKIPS", skip), rfc.Int(handle, "ROWCOUNT", pageSize)
        fields := rfc.Table(handle, "FIELDS")
        for name in fieldNames
            rfc.Set(rfc.Append(fields), "FIELDNAME", name)
        opts := rfc.Table(handle, "OPTIONS")
        for line in options {
            if StrLen(line) > 72
                throw Error("Query line too long")
            rfc.Set(rfc.Append(opts), "TEXT", line)
        }
        rfc.Invoke(handle)
        layout := []
        Loop rfc.Count(fields) {
            rfc.Move(fields, A_Index - 1)
            layout.Push(Map("name", rfc.Get(fields, "FIELDNAME"), "offset", Integer(rfc.Get(fields, "OFFSET")), "length", Integer(rfc.Get(fields, "LENGTH"))))
        }
        data := rfc.Table(handle, "DATA"), count := rfc.Count(data)
        Loop count {
            rfc.Move(data, A_Index - 1), row := Map()
            ; Get()의 Trim은 선행 공백을 제거할 수 있으므로 WA는 직접 읽는다.
            memory := Buffer(4096, 0), length := 0
            rfc.Check(DllCall(rfc.Fn("RfcGetString"), "Ptr", data, "Str", "WA", "Ptr", memory, "UInt", 2047, "UIntP", &length, "Ptr", rfc.err, "Int"), "Read row")
            text := StrGet(memory, length, "UTF-16")
            for field in layout
                row[field["name"]] := Trim(SubStr(text, field["offset"] + 1, field["length"]))
            rows.Push(row)
            if rows.Length > maxRows
                throw Error("Result exceeds limit")
        }
        if count < pageSize
            break
        skip += count
    }
    return rows
}

SapMaterialOptions(materials) {
    options := ["("]
    for index, material in materials {
        if !RegExMatch(material, "^[0-9A-Za-z_-]{1,18}$")
            throw Error("Unsupported material number")
        options.Push((index = 1 ? "" : "OR ") "MATNR = '" material "'")
    }
    options.Push(")")
    return options
}

QueryMaterialLocations(rfc, locations, searchText) {
    allowed := Map()
    for code in ["5111", "5112", "5113", "5121", "5122", "5123", "5124", "5125", "5126", "5127", "5128", "5129", "5130", "5131", "5132", "5133"]
        allowed[code] := true
    if !(locations is Array) || !locations.Length || locations.Length > allowed.Count || StrLen(searchText) > 200
        throw Error("Invalid query")
    stocks := [], uniqueMaterials := Map(), seenLocations := Map(), byLocation := Map()
    for code in locations {
        if !allowed.Has(code) || seenLocations.Has(code)
            throw Error("Invalid storage location")
        seenLocations[code] := true
        localRows := SapReadTable(rfc, "MARD", ["MATNR", "WERKS", "LGORT", "LGPBE", "LABST"], ["WERKS = '5010'", "AND LGORT = '" code "'"])
        byLocation[code] := localRows
        for row in localRows {
            uniqueMaterials[row["MATNR"]] := true
            stocks.Push(row)
        }
        if stocks.Length > 30000
            throw Error("Too many results")
    }
    descriptions := Map(), batches := Map(), materialChunks := [], chunk := []
    for material in uniqueMaterials {
        chunk.Push(material)
        if chunk.Length = 35 {
            materialChunks.Push(chunk), chunk := []
        }
    }
    if chunk.Length
        materialChunks.Push(chunk)
    for chunk in materialChunks {
        opts := SapMaterialOptions(chunk)
        opts.Push("AND SPRAS = '3'")
        for row in SapReadTable(rfc, "MAKT", ["MATNR", "MAKTX"], opts)
            descriptions[row["MATNR"]] := row["MAKTX"]
    }
    matches := Map()
    for material, description in descriptions {
        if SapMaterialMatches(description, searchText)
            matches[material] := true
    }
    if !matches.Count
        return []
    materialChunks := [], chunk := []
    for material in matches {
        chunk.Push(material)
        if chunk.Length = 35
            materialChunks.Push(chunk), chunk := []
    }
    if chunk.Length
        materialChunks.Push(chunk)
    ; 배치명만으로 신품/중고품을 추측하지 않고 실제 평가유형을 읽는다.
    for chunk in materialChunks {
        opts := SapMaterialOptions(chunk), opts.Push("AND WERKS = '5010'")
        for row in SapReadTable(rfc, "MCHA", ["MATNR", "CHARG", "BWTAR"], opts)
            batches[row["MATNR"] "|" row["CHARG"]] := row["BWTAR"]
    }
    quantities := Map()
    for code in locations {
        for chunk in materialChunks {
            opts := SapMaterialOptions(chunk), opts.Push("AND WERKS = '5010'"), opts.Push("AND LGORT = '" code "'")
            for row in SapReadTable(rfc, "MCHB", ["MATNR", "CHARG", "CLABS"], opts) {
                batchKey := row["MATNR"] "|" row["CHARG"], key := row["MATNR"] "|" code
                if !quantities.Has(key)
                    quantities[key] := Map("new", 0, "used", 0, "total", 0)
                value := Number(row["CLABS"]), quantities[key]["total"] += value
                if !batches.Has(batchKey) && value != 0
                    throw Error("Unknown batch valuation")
                kind := batches.Has(batchKey) ? batches[batchKey] : ""
                if kind = "신품"
                    quantities[key]["new"] += value
                else if kind = "중고품"
                    quantities[key]["used"] += value
            }
        }
    }
    result := [], seen := Map()
    for row in stocks {
        material := row["MATNR"], code := row["LGORT"], key := material "|" code
        if !matches.Has(material)
            continue
        if seen.Has(key)
            throw Error("Duplicate material row")
        seen[key] := true
        qty := quantities.Has(key) ? quantities[key] : Map("new", 0, "used", 0, "total", 0)
        if Abs(qty["total"] - Number(row["LABST"])) > 0.0005
            throw Error("Stock totals changed or unsupported nonbatch stock")
        result.Push(Map("material", material, "description", descriptions[material], "location", code, "bin", row["LGPBE"], "newStock", qty["new"], "usedStock", qty["used"]))
    }
    return result
}

SapMaterialMatches(description, searchText) {
    tokens := StrSplit(RegExReplace(Trim(searchText), "\s+", " "), " ")
    for token in tokens
        if token != "" && !InStr(description, token, false)
            return false
    return true
}
