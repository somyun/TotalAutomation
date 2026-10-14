class SapRfcClient {
    __New() {
        this.err := Buffer(8192, 0)
        this.module := DllCall("LoadLibraryW", "Str", "C:\Windows\SysWOW64\sapnwrfc.dll", "Ptr")
        if !this.module
            throw OSError(A_LastError, "SAP RFC library could not be loaded")
        this.funcs := Map()
        this.conn := 0
        this.handles := []
    }
    Fn(name) {
        if !this.funcs.Has(name) {
            ptr := DllCall("GetProcAddress", "Ptr", this.module, "AStr", name, "Ptr")
            if !ptr
                throw Error("Missing SDK function: " name)
            this.funcs[name] := ptr
        }
        return this.funcs[name]
    }
    Check(rc, stage) {
        if rc != 0
            throw Error(stage ": " StrGet(this.err.Ptr + 8, "UTF-16") " / " StrGet(this.err.Ptr + 264, "UTF-16"))
    }
    Open(values) {
        memory := Buffer(values.Count * 2 * A_PtrSize, 0), strings := [], index := 0
        for key, value in values {
            keyBuf := Buffer(StrPut(key, "UTF-16") * 2, 0), valBuf := Buffer(StrPut(value, "UTF-16") * 2, 0)
            StrPut(key, keyBuf, "UTF-16"), StrPut(value, valBuf, "UTF-16")
            strings.Push(keyBuf, valBuf)
            NumPut("Ptr", keyBuf.Ptr, "Ptr", valBuf.Ptr, memory, index * 2 * A_PtrSize)
            index++
        }
        this.conn := DllCall(this.Fn("RfcOpenConnection"), "Ptr", memory, "UInt", values.Count, "Ptr", this.err, "Ptr")
        if !this.conn
            this.Check(1, "RFC logon")
    }
    Function(name) {
        desc := DllCall(this.Fn("RfcGetFunctionDesc"), "Ptr", this.conn, "Str", name, "Ptr", this.err, "Ptr")
        if !desc
            this.Check(1, "Function metadata " name)
        handle := DllCall(this.Fn("RfcCreateFunction"), "Ptr", desc, "Ptr", this.err, "Ptr")
        if !handle
            this.Check(1, "Create " name)
        this.handles.Push(handle)
        return handle
    }
    Invoke(handle) => this.Check(DllCall(this.Fn("RfcInvoke"), "Ptr", this.conn, "Ptr", handle, "Ptr", this.err, "Int"), "Invoke")
    Structure(handle, name) {
        ptr := 0
        this.Check(DllCall(this.Fn("RfcGetStructure"), "Ptr", handle, "Str", name, "PtrP", &ptr, "Ptr", this.err, "Int"), "Structure " name)
        return ptr
    }
    Table(handle, name) {
        ptr := 0
        this.Check(DllCall(this.Fn("RfcGetTable"), "Ptr", handle, "Str", name, "PtrP", &ptr, "Ptr", this.err, "Int"), "Table " name)
        return ptr
    }
    Append(table) {
        ptr := DllCall(this.Fn("RfcAppendNewRow"), "Ptr", table, "Ptr", this.err, "Ptr")
        if !ptr
            this.Check(1, "Append range")
        return ptr
    }
    Count(table) {
        count := 0
        this.Check(DllCall(this.Fn("RfcGetRowCount"), "Ptr", table, "UIntP", &count, "Ptr", this.err, "Int"), "Row count")
        return count
    }
    Move(table, index) => this.Check(DllCall(this.Fn("RfcMoveTo"), "Ptr", table, "UInt", index, "Ptr", this.err, "Int"), "Move row")
    Set(handle, name, value) => this.Check(DllCall(this.Fn("RfcSetChars"), "Ptr", handle, "Str", name, "Str", value, "UInt", StrLen(value), "Ptr", this.err, "Int"), "Set " name)
    Int(handle, name, value) => this.Check(DllCall(this.Fn("RfcSetInt"), "Ptr", handle, "Str", name, "Int", value, "Ptr", this.err, "Int"), "Set " name)
    Get(handle, name) {
        memory := Buffer(4096, 0), length := 0
        this.Check(DllCall(this.Fn("RfcGetString"), "Ptr", handle, "Str", name, "Ptr", memory, "UInt", 2047, "UIntP", &length, "Ptr", this.err, "Int"), "Get " name)
        return Trim(StrGet(memory, length, "UTF-16"))
    }
    Fields(handle) {
        type := DllCall(this.Fn("RfcDescribeType"), "Ptr", handle, "Ptr", this.err, "Ptr"), count := 0, names := []
        if !type
            this.Check(1, "Describe type")
        this.Check(DllCall(this.Fn("RfcGetFieldCount"), "Ptr", type, "UIntP", &count, "Ptr", this.err, "Int"), "Field count")
        desc := Buffer(8192, 0)
        Loop count {
            this.Check(DllCall(this.Fn("RfcGetFieldDescByIndex"), "Ptr", type, "UInt", A_Index - 1, "Ptr", desc, "Ptr", this.err, "Int"), "Field metadata")
            names.Push(StrGet(desc.Ptr, "UTF-16"))
        }
        return names
    }
    Returns(handle) {
        table := this.Table(handle, "RETURN"), messages := []
        Loop this.Count(table) {
            this.Move(table, A_Index - 1)
            message := Map()
            for field in ["TYPE", "ID", "NUMBER", "MESSAGE"]
                message[field] := this.Get(table, field)
            messages.Push(message)
            if message["TYPE"] = "E" || message["TYPE"] = "A" || message["TYPE"] = "X"
                throw Error("BAPI returned " message["TYPE"] ": " message["MESSAGE"])
        }
        return messages
    }
    Close() {
        for handle in this.handles
            DllCall(this.Fn("RfcDestroyFunction"), "Ptr", handle, "Ptr", this.err, "Int")
        this.handles := []
        if this.conn
            DllCall(this.Fn("RfcCloseConnection"), "Ptr", this.conn, "Ptr", this.err, "Int")
        this.conn := 0
        if this.module
            DllCall("FreeLibrary", "Ptr", this.module)
        this.module := 0
    }
}
LoadSession() {
    encrypted := FileRead(A_ScriptDir "\.session.dat", "RAW")
    input := Buffer(8, 0), output := Buffer(8, 0)
    NumPut("UInt", encrypted.Size, "Ptr", encrypted.Ptr, input)
    if !DllCall("Crypt32\CryptUnprotectData", "Ptr", input, "Ptr", 0, "Ptr", 0, "Ptr", 0, "Ptr", 0, "UInt", 1, "Ptr", output, "Int")
        throw OSError(A_LastError, "Session decrypt")
    ptr := NumGet(output, 4, "Ptr")
    try return JSON.parse(StrGet(ptr, NumGet(output, 0, "UInt"), "UTF-8"))
    finally DllCall("LocalFree", "Ptr", ptr)
}

AddRange(rfc, table, name, value, option := "EQ", high := "") {
    row := rfc.Append(table)
    for field, data in Map("FIELD_NAME", name, "SIGN", "I", "OPTION", option, "LOW_VALUE", value, "HIGH_VALUE", high)
        rfc.Set(row, field, data)
}
