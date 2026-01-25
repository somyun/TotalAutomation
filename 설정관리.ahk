
class ConfigManager {
    static ConfigPath := A_ScriptDir "\config.json"
    static Config := Map()
    static CurrentUser := Map() ; 현재 로그인한 유저의 '프로필' 객체 (user.profile)
    static NeedsPasswordReset := false ; 복호화 실패로 인한 암호 초기화 발생 플래그

    ; 설정 로드
    static Load() {
        if !FileExist(this.ConfigPath) {
            this.Config := Map("users", Map(), "appSettings", Map())
            this.Save() ; 즉시 파일 생성
            return true
        }

        try {
            fileContent := FileRead(this.ConfigPath, "UTF-8")
            this.Config := JSON.parse(fileContent)

            ; 플래그 초기화
            this.NeedsPasswordReset := false
            hasDecryptionError := false

            ; 민감한 필드 복호화
            key := this.GetKey()
            if (key != "" && this.Config.Has("users")) {
                for uid, user in this.Config["users"] {
                    if user.Has("profile") {
                        ; _DecryptProfile이 실패(false)를 반환하면 에러 플래그 설정
                        if !this._DecryptProfile(user["profile"], key) {
                            hasDecryptionError := true
                        }
                    }
                }
            }

            ; 복호화 오류가 하나라도 있었다면
            if (hasDecryptionError) {
                this.NeedsPasswordReset := true
                this.Save() ; 빈 암호를 파일에 즉시 저장
            }

            return true

        } catch as e {
            MsgBox("설정 파일 로드 중 오류 발생: " e.Message, "오류", "Iconx")
            return false
        }
    }

    ; 설정 저장
    static Save() {
        try {
            ; 1. 파일이 존재하는지 확인 (파일 ID를 얻기 위해)
            if !FileExist(this.ConfigPath)
                FileAppend("", this.ConfigPath, "UTF-8")

            ; 2. 저장할 데이터 준비 (민감한 필드 암호화)
            key := this.GetKey()

            ; 메모리 내 구성을 수정하지 않도록 JSON을 통한 깊은 복사
            configToSave := JSON.parse(JSON.stringify(this.Config))

            if (key != "" && configToSave.Has("users")) {
                for uid, user in configToSave["users"] {
                    if user.Has("profile") {
                        this._EncryptProfile(user["profile"], key)
                    }
                }
            }

            fileContent := JSON.stringify(configToSave, 4)

            ; 3. 파일 덮어쓰기 (FileOpen을 "w"로 사용하여 파일 ID 보존)
            fObj := FileOpen(this.ConfigPath, "w", "UTF-8")
            fObj.Write(fileContent)
            fObj.Close()

            return true
        } catch as e {
            MsgBox("설정 파일 저장 중 오류 발생: " e.Message, "오류", "Iconx")
            return false
        }
    }

    ; 전역 설정 가져오기 (appSettings 등)
    static Get(keyPath, defaultValue := "") {
        return this._GetValue(this.Config, keyPath, defaultValue)
    }

    ; 전역 설정 설정하기
    static Set(keyPath, value) {
        this._SetValue(this.Config, keyPath, value)
        this.Save()
    }

    ; --- 사용자별 설정 메서드 ---

    ; 현재 로그인한 유저의 ID 가져오기
    static GetCurrentUserID() {
        if this.CurrentUser.Has("id")
            return this.CurrentUser["id"]
        return ""
    }

    ; 특정 유저의 루트 객체 가져오기
    static GetUserRoot(userID) {
        if !this.Config.Has("users")
            this.Config["users"] := Map()

        if !this.Config["users"].Has(userID)
            this._InitializeUser(userID)

        return this.Config["users"][userID]
    }

    ; 현재 유저의 설정 가져오기 (예: "hotkeys", "presets.dailyLog")
    static GetUserConfig(keyPath, defaultValue := "") {
        uid := this.GetCurrentUserID()
        if (uid == "")
            return defaultValue
        userRoot := this.GetUserRoot(uid)
        return this._GetValue(userRoot, keyPath, defaultValue)
    }

    ; 현재 유저의 설정 저장하기
    static SetUserConfig(keyPath, value) {
        uid := this.GetCurrentUserID()
        if (uid == "")
            return
        userRoot := this.GetUserRoot(uid)
        this._SetValue(userRoot, keyPath, value)
        this.Save()
    }

    ; 유저 초기화 (기본 구조 생성)
    static _InitializeUser(userID) {
        this.Config["users"][userID] := Map(
            "profile", Map("id", userID),
            "hotkeys", [],
            "presets", Map()
        )
    }

    ; 프로필 목록 가져오기 (로그인 화면용 - 모든 유저의 profile 객체 리스트 반환)
    static GetProfiles() {
        profiles := []
        if this.Config.Has("users") {
            for uid, data in this.Config["users"] {
                if data.Has("profile")
                    profiles.Push(data["profile"])
            }
        }
        return profiles
    }

    ; 내부 헬퍼: Get Value
    static _GetValue(rootObj, keyPath, defaultValue) {
        keys := StrSplit(keyPath, ".")
        current := rootObj

        for k in keys {
            if !IsObject(current) || !current.Has(k)
                return defaultValue
            current := current[k]
        }
        return current
    }

    ; 내부 헬퍼: Set Value
    static _SetValue(rootObj, keyPath, value) {
        keys := StrSplit(keyPath, ".")
        current := rootObj

        loop keys.Length - 1 {
            k := keys[A_Index]
            if !current.Has(k)
                current[k] := Map()
            current := current[k]
        }

        lastKey := keys[keys.Length]
        current[lastKey] := value
    }
    static GetKey() {
        if FileExist(this.ConfigPath)
            return GetFileID(this.ConfigPath)
        return ""
    }

    static _EncryptProfile(profile, key) {
        fields := ["pw2", "sapPW", "webPW"]
        for field in fields {
            if profile.Has(field) && profile[field] != "" {
                ; 사본에서 작업하지만 문제가 발생할 경우 이중 암호화 방지.
                ; 이미 암호화되었는지 확인용 첨두어 "ENC_"
                if (SubStr(profile[field], 1, 4) != "ENC_") {
                    ;데이터 앞에 "_OK_" 검증용 문자열을 붙여서 암호화
                    profile[field] := "ENC_" . Encrypt("_OK_" profile[field], key)
                }
            }
        }
    }

    static _DecryptProfile(profile, key) {
        fields := ["pw2", "sapPW", "webPW"]
        allSuccess := true ; 전체 성공 여부

        for field in fields {
            if profile.Has(field) && profile[field] != "" {
                if (SubStr(profile[field], 1, 4) == "ENC_") {
                    try {
                        decryptedRaw := Decrypt(SubStr(profile[field], 5), key)

                        ; 복호화된 데이터가 "_OK_"로 시작하는지 검사
                        if (SubStr(decryptedRaw, 1, 4) == "_OK_") {
                            ; 검증 성공: 앞의 "_OK_" 4글자를 떼고 실제 데이터만 저장
                            profile[field] := SubStr(decryptedRaw, 5)
                        } else {
                            ; 검증 실패: 강제로 에러 발생시킴
                            throw Error("Decryption Validation Failed")
                        }

                    } catch {
                        ; 이제 여기서 확실하게 잡힙니다.
                        profile[field] := ""
                        allSuccess := false
                    }
                }
            }
        }
        return allSuccess ; 성공하면 true, 하나라도 실패하면 false 반환
    }
}
