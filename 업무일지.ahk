#Requires AutoHotkey v2.0

class WorkLogManager {
    static BaseDate := "20260101"   ; 기준일

    ; 주간 근무 순서: A -> D -> C -> B
    static DayShiftOrder := ["A", "D", "C", "B"]

    ; 야간 근무 순서: B -> A -> D -> C
    static NightShiftOrder := ["B", "A", "D", "C"]

    ; --------------------------------------------------------------------------
    ; 현재 컨텍스트(이전/현재/다음 근무) 반환
    ; --------------------------------------------------------------------------
    static GetCurrentContext(userTeam := "") {
        now := A_Now

        ; 00:00 ~ 09:00 사이인 경우 전날 야간으로 간주 (전역 로직)
        hour := Integer(FormatTime(now, "H"))
        min := Integer(FormatTime(now, "m"))
        timeVal := hour * 100 + min ; 비교를 쉽게 하기 위한 HHmm 형식

        isNightCorrection := (hour < 9)

        targetDate := now
        if (isNightCorrection) {
            targetDate := DateAdd(now, -1, "Days")
        }

        ; 전역 표준 로직
        ; 09:00 ~ 17:59 : 주간
        ; 18:00 ~ 08:59 : 야간 (익일 09시 전까지)
        isNight := (hour >= 18 || hour < 9)

        currentTeam := this.CalculateTeam(targetDate, isNight)

        ; --- 버퍼 로직 재정의 (출근 30분) ---
        ; 아침 버퍼: 08:30 (830) ~ 09:00 (900)
        ; 저녁 버퍼: 17:30 (1730) ~ 18:00 (1800)
        if (userTeam != "") {
            ; 조 이름 정규화 ("조" 접미사가 있다면 제거)
            userTeam := StrReplace(userTeam, "조", "")

            ; 1. 아침 버퍼 (08:30 ~ 09:00)
            if (timeVal >= 830 && timeVal <= 900) {
                ; 주간 (출근)
                teamDay := this.CalculateTeam(now, false) ; 오늘 주간

                if (userTeam == teamDay) {
                    ; 주간 컨텍스트 강제
                    isNight := false
                    currentTeam := teamDay
                }
            }
            ; 2. 저녁 버퍼 (17:30 ~ 18:00)
            else if (timeVal >= 1730 && timeVal <= 1800) {
                ; 야간 (출근)
                teamNight := this.CalculateTeam(now, true) ; 오늘 야간

                if (userTeam == teamNight) {
                    ; 야간 컨텍스트 강제
                    isNight := true
                    currentTeam := teamNight
                }
            }
        }

        ; 이전/다음 근무 계산 (확정된 isNight/currentTeam 기준)
        if (!isNight) { ; 현재 주간
            prevTeam := this.CalculateTeam(DateAdd(targetDate, -1, "Days"), true) ; 전날 야간
            nextTeam := this.CalculateTeam(targetDate, true)                      ; 오늘 야간

            shiftName := "주간"
        } else { ; 현재 야간
            prevTeam := this.CalculateTeam(targetDate, false)                     ; 오늘 주간
            nextTeam := this.CalculateTeam(DateAdd(targetDate, 1, "Days"), false) ; 내일 주간

            shiftName := "야간"
        }

        return Map(
            "date", FormatTime(targetDate, "yyyyMMdd"),
            "isNight", isNight,
            "shiftName", shiftName,
            "prev", prevTeam,
            "current", currentTeam,
            "next", nextTeam
        )
    }

    ; --------------------------------------------------------------------------
    ; 특정 날짜/시간대의 근무조 계산
    ; --------------------------------------------------------------------------
    static CalculateTeam(dateStr, isNight) {
        ; 날짜 차이 계산 (일 단위)
        diff := DateDiff(dateStr, this.BaseDate, "Days")

        ; 음수 처리 (기준일 이전)
        if (diff < 0) {
            diff := Mod(diff, 4) + 4
        }

        idx := Mod(diff, 4) + 1 ; 1부터 시작하는 인덱스

        if (isNight) {
            return this.NightShiftOrder[idx]
        } else {
            return this.DayShiftOrder[idx]
        }
    }
}

class workers {
    ; 클래스의 속성을 정의합니다.
    Day := true
    Weekday := true
    schedule := ""

    ;일지생성 야간 작성/승인자
    nightwriter := ""               ;야간작성자
    nightapprover := ""             ;야간승인자

    ;인수자
    handover := ""                  ;인수자

    ;교육사항 인원
    TotalCnt := 0					;총원
    attendCnt := 0					;참석자
    absentCnt := 0					;결석자
    tutor := ""						;강사
    _vacaReason := Map()			;불참사유

    ;기타사항 인원(음주검사)
    firstChecker := ""				;음주검사자
    secondChecker := ""				;음주검사자 차석
    drinkAttendees := ""			;음주측정대상자

    ;운전적합성검사 인원
    drivers := array()				;모터카운전원들
    driverChker := ""				;운전적합성 검사자
    TrackSafetyManager := ""		;철도운행안전관리자 - 야간, 운전원을 제외한 인원 중 차선임자, 4명 이상 근무 시에만 선임

    ;안전관리 인원
    firstWorker := ""				;작업책임자
    workers := ""					;작업자
    workersArray := Array()			;작업자 배열

    ; 생성자: app.js에서 전달받은 data 객체를 직접 파싱하여 초기화합니다.
    __New(data) {
        ; 1. 스케줄 설정 (app.js: workType="day"/"night")
        this.schedule := (data.Has("workType") && data["workType"] == "day") ? "주간" : "야간"
        this.day := (this.schedule == "주간")
        this.weekday := (this.schedule == "주간")

        ; 2. 작업자 및 운전원 초기화
        workerList := data.Has("workers") ? data["workers"] : []
        this.drivers := []

        for _worker in workerList {
            name := _worker["name"]
            ; app.js: reason(휴가/비고)와 team(근무조) 분리 수신
            reason := _worker.Has("reason") ? _worker["reason"] : ""
            team := _worker.Has("team") ? _worker["team"] : ""
            attend := _worker["attend"]

            ; driverRole 처리 ("정", "부", "검사자")
            if (_worker.Has("driverRole") && _worker["driverRole"] != "") {
                role := _worker["driverRole"]
                if (role == "정")
                    this.drivers.InsertAt(1, name)
                else if (role == "부")
                    this.drivers.Push(name)
                else if (role == "검사자")
                    this.driverChker := name ; 명시적 검사자 지정
            }

            if !attend {
                if !((name == "분소장" or team == "일근") and !this.weekday) {
                    this.absentCnt++
                    if (reason != "") { ; 사유가 있을 때만 집계
                        if !this._vacaReason.Has(reason)
                            this._vacaReason[reason] := ""
                        this._vacaReason[reason] .= name ", "
                    }
                }
            }
            else {	;출근자
                this.workersArray.Push(name)
                ;음주
                if this.firstChecker == ""
                    this.firstChecker := name			;측정자가 비어있으면 지정
                else {
                    if this.firstChecker != "" and this.secondChecker == ""
                        this.secondChecker := name		;차선임 측정자가 비어있으면 지정
                    this.drinkAttendees .= name ", "	;측정대상자 누적
                }

                ;작업 전 교육
                if this.tutor == ""
                    this.tutor := name					;최선임자가 강사
                this.attendCnt++						;참석자 체크

                ;안전
                if name != "분소장" and this.firstWorker == ""		;최선임자가
                    this.firstWorker := name						;조책임자
                else
                    if !(name == "분소장" or team == "일근") ; 일근자 제외
                        this.workers .= name ", "

                ;운전적합성 검사자 (자동 지정 로직 - 명시적 지정이 없을 때 동작)
                if this.driverChker == "" and this.drivers.Length {	;검사자가 비어있고 운전자가 있고
                    yesDriver := false
                    for drivername in this.drivers
                        if drivername == name
                            yesDriver := true
                    if !yesDriver						;운전자가 아니면 지정
                        this.driverChker := name
                }

                ;철운안자
                ; 야간, 운전자를 제외한 차석, 아직 철운안자 미 지정 시
                try {
                    if this.day or this.drivers.Length > 0 and this.drivers[1] == name or StrLen(this.TrackSafetyManager
                    ) > 1
                        continue
                    else
                        if this.TrackSafetyManager == ""
                            this.TrackSafetyManager := 1
                        else
                            this.TrackSafetyManager := name
                }

            }

        }	;인원파악 끝

        this.TotalCnt := this.attendCnt + this.absentCnt
        this.workers := SubStr(this.workers, 1, -2)
        this.drinkAttendees := SubStr(this.drinkAttendees, 1, -2)

        ;철운안자 4명 이상 근무 시에만 선임
        if this.attendCnt < 4
            this.TrackSafetyManager := ""

        ; --------------------------------------------------------------------------
        ; [New] 인수자 및 야간 근무자(작성/승인) 자동 계산 로직
        ; --------------------------------------------------------------------------
        try {
            ; 1. 현재 유저의 근무조 확인
            currentUserTeam := ""
            if (ConfigManager.CurrentUser.Has("team")) {
                currentUserTeam := ConfigManager.CurrentUser["team"]
            }

            if (currentUserTeam != "") {
                ; 2. 다음 근무조 계산 (WorkLogManager 활용)
                ; 주의: 현재 컨텍스트가 아닌 '나의 근무조' 기준의 다음 주기를 찾아야 하므로
                ; WorkLogManager.GetCurrentContext(currentUserTeam)을 호출하면
                ; 현재 시간과 내 근무조에 맞는 'Next' 팀을 반환함.
                context := WorkLogManager.GetCurrentContext(currentUserTeam)
                nextTeam := context["next"]

                ; 3. 전체 직원 명단에서 다음 근무조 인원 필터링
                allColleagues := ConfigManager.Get("appSettings")["colleagues"]
                nextTeamMembers := []

                for col in allColleagues {
                    if (col.Has("team") && col["team"] == nextTeam) {
                        nextTeamMembers.Push(col)
                    }
                }

                ; 4. 사번 순 정렬 (오름차순: 낮은 숫자가 선임)
                ; Bubble Sort
                if (nextTeamMembers.Length > 1) {
                    loop nextTeamMembers.Length - 1 {
                        i := A_Index
                        loop nextTeamMembers.Length - i {
                            j := A_Index
                            id1 := Number(nextTeamMembers[j]["id"])
                            id2 := Number(nextTeamMembers[j + 1]["id"])
                            if (id1 > id2) {
                                temp := nextTeamMembers[j]
                                nextTeamMembers[j] := nextTeamMembers[j + 1]
                                nextTeamMembers[j + 1] := temp
                            }
                        }
                    }
                }

                if (nextTeamMembers.Length > 0) {
                    senior := nextTeamMembers[1]["name"]
                    junior := nextTeamMembers[nextTeamMembers.Length]["name"]

                    ; 5. 인수자 (항상 다음 근무조 최선임)
                    this.handover := senior

                    ; 6. 야간 작성자/승인자 (주간 근무일 때만 필요)
                    if (this.day) {
                        this.nightwriter := junior  ; 야간 최후임
                        this.nightapprover := senior ; 야간 최선임
                    }
                }
            }
        } catch as e {
            ; 오류 발생 시 기본값 유지 (공란)
            ; MsgBox("인수자 계산 중 오류: " e.Message)
        }
    }	;생성자 끝

    getReasons() {
        result := ""
        for reason, names in this._vacaReason {
            result .= reason "(" SubStr(names, 1, -2) "), "
        }
        return SubStr(result, 1, -2)
    }

}

class RunWorkLog {
    cUIA := ""

    __New(data) {
        this.data := data

        ; 1. 옵션 매핑 (app.js options{...} -> AHK this.data["chk..."])
        if (data.Has("options")) {
            opt := data["options"]
            this.data["chkCreate"] := opt.Has("makeLog") ? opt["makeLog"] : false
            this.data["chkGeneralWork"] := opt.Has("general") ? opt["general"] : false
            this.data["chkSafety"] := opt.Has("safe") ? opt["safe"] : false
            this.data["chkDriving"] := opt.Has("driving") ? opt["driving"] : false
            this.data["chkDrinkDetect"] := opt.Has("drink") ? opt["drink"] : false
        }

        ; 2. 안전관리 데이터 변환 (app.js safetyData{...} -> AHK this.data["safetyList"][...])
        if (data.Has("safetyData")) {
            sData := data["safetyData"]
            safetyList := []

            ; Row 1
            if (sData.Has("content1") && sData["content1"] != "") {
                item := Map()
                item["content"] := sData["content1"]

                ; 날짜+시간(HH:mm -> HHmm) 병합
                d := sData.Has("date1") ? sData["date1"] : ""
                s := StrReplace(sData.Has("start1") ? sData["start1"] : "", ":", "")
                e := StrReplace(sData.Has("end1") ? sData["end1"] : "", ":", "")

                item["start"] := d . s
                item["end"] := d . e
                ; confirm 등은 기본값 사용(allOK)

                safetyList.Push(item)
            }

            ; Row 2
            if (sData.Has("content2") && sData["content2"] != "") {
                item := Map()
                item["content"] := sData["content2"]

                d := sData.Has("date2") ? sData["date2"] : ""
                s := StrReplace(sData.Has("start2") ? sData["start2"] : "", ":", "")
                e := StrReplace(sData.Has("end2") ? sData["end2"] : "", ":", "")

                item["start"] := d . s
                item["end"] := d . e

                safetyList.Push(item)
            }

            this.data["safetyList"] := safetyList
        }

        this.workers := workers(data)
    }

    Run() {
        ; 1. 일지 생성
        if this.data["chkCreate"] {
            if this.EnsureReady("WorkLog_Create")
                this.CreateLog()
            else {
                MsgBox("일지 생성 준비에 실패하였습니다.", "알림", "IconStop")
                return false
            }
        }

        ; 2. 일반 업무
        if this.data["chkGeneralWork"] {
            if this.cUIA or this.EnsureReady("WorkLog_View")
                this.GeneralWork()
            else {
                MsgBox("일지 조회 준비에 실패하였습니다.", "알림", "IconStop")
                return false
            }
        }

        ; 3. 안전 관리
        if this.data["chkSafety"] {
            if this.cUIA or this.EnsureReady("WorkLog_View")
                this.SafeManage()
            else {
                MsgBox("일지 조회 준비에 실패하였습니다.", "알림", "IconStop")
                return false
            }
        }

        ; 4. 운전 적합성
        if this.data["chkDriving"] {
            if this.cUIA or this.EnsureReady("WorkLog_View")
                this.DrivingCheck()
            else {
                MsgBox("일지 조회 준비에 실패하였습니다.", "알림", "IconStop")
                return false
            }
        }

        return true
    }

    EnsureReady(type) {
        this.cUIA := WebAutoLogin.EnsureReady(type)
        if (!this.cUIA) {
            StopMacro()
            return false
        }
        return true
    }

    CreateLog() {
        approver := this.data.Has("approver") ? this.data["approver"] : "이윤수"
        nightMaker := this.data.Has("nightMaker") ? this.data["nightMaker"] : "황혁수"
        nightChecker := this.data.Has("nightChecker") ? this.data["nightChecker"] : "조구형"

        ;근무방식, 기존 컨트롤클릭
        ;this.cUIA.WaitElement({ AutomationId: "I_GYELJAE" }, 3000).ControlClick("left")
        ;this.cUIA.send "{down}{enter}"

        comboBox := this.cUIA.WaitElement({ AutomationId: "I_GYELJAE" }, 3000)
        comboBox.expand()
        Sleep 50
        comboBox.WaitElement({ Name: "교대" }, 3000).invoke()
        Sleep 100
        comboBox.collapse()

        ;검토자
        일근 := this.cUIA.WaitElement({ AutomationId: "pernrView1" }, 3000)
        if 일근.WaitElement({ AutomationId: "I_PERNR12_TEXT" }).value != this.workers.firstWorker {
            일근.WaitElement({ Type: "Image", Name: "검토자" }).Invoke()
            lensInput(this.cUIA.browserID, , this.workers.firstWorker)	;부서, 이름 검색 후 클릭
        }

        ;승인자
        WinWaitActive(this.cUIA.browserID)
        if 일근.WaitElement({ AutomationId: "I_PERNR13_TEXT" }).value != approver {
            일근.WaitElement({ Type: "Image", Name: "승인자" }).Invoke()
            lensInput(this.cUIA.browserID, , approver)
        }

        ;야간
        야간 := this.cUIA.WaitElement({ AutomationId: "pernrView2" })
        ;작성자
        WinWaitActive(this.cUIA.browserID)
        야간.WaitElement({ Type: "Image", Name: "작성자" }).Invoke()
        lensInput(this.cUIA.browserID, , nightMaker)

        ;검토자
        WinWaitActive(this.cUIA.browserID)
        야간.WaitElement({ Type: "Image", Name: "검토자" }).Invoke()
        lensInput(this.cUIA.browserID, , nightChecker)

        ;승인자
        WinWaitActive(this.cUIA.browserID)
        if 야간.WaitElement({ AutomationId: "I_PERNR23_TEXT" }).value != approver {
            야간.WaitElement({ Type: "Image", Name: "승인자" }).Invoke()
            lensInput(this.cUIA.browserID, , approver)
        }

        if MsgBox("기본정보저장을 하시겠습니까?", , "Y/N icon?") == "Yes" {
            this.cUIA.FindElement({ LocalizedType: "링크", Name: "기본정보저장" }).Invoke()
            sleep 1000

            this.cUIA.WaitElement({ AutomationId: "btnOpn" }, 3000).Invoke()
            table := this.cUIA.FindElement({ AutomationId: "tb_person" })
            table.WaitElement({ Name: approver }, 3000) ;approver
            Sleep 250

            this.cUIA.FindElement({ AutomationId: "btnSave" }).Invoke()
            this.cUIA.WaitElement({ LocalizedType: "단추", Name: "확인" }, 3000).Invoke()
            this.cUIA.WaitElement({ LocalizedType: "단추", Name: "확인" }, 3000).Invoke()
        }
        else
            return false
    }

    GeneralWork() {
        safeContents := ["안전 확보 요강 및 기타 작업에 필요한 사항 등", "안전보건 11대 수칙", "전기원 안전수칙", "일반전기작업 안전수칙", "시설물 안전점검 수칙",
            "작업장 내 정리 정돈 안전 수칙", "사다리 작업 및 정전작업 요령"]

        menuClick(this.cUIA, "일반업무")							;메뉴진입

        ;주간
        if InStr(this.workers.schedule, "주간") {
            ;인수인계
            table := this.cUIA.WaitElement({ AutomationId: "tb1" }, 2000)
            Sleep 250
            table.FindElement({ AutomationId: "I_INGYEP" }).value := this.workers.firstworker
            table.FindElement({ AutomationId: "I_INSUP" }).value := this.workers.handover
            table.FindElement({ AutomationId: "I_ITEXT" }).value := "특이사항 없음."

            ;안전의날(매주금요일)/안전점검의날(매월4일)
            if A_DD == 4 or A_DDD == "금" {
                this.cUIA.FindElement({ AutomationId: "btnOpen_05" }).Invoke()
            }
            this.cUIA.FindElement({ AutomationId: "btnOpen_06" }).Invoke()	;에너지절약 점검표
            this.cUIA.FindElement({ AutomationId: "btnOpen_07" }).Invoke()	;안전보건 11대 수칙
            this.cUIA.FindElement({ AutomationId: "btnOpen_11" }).Invoke()	;작업 전 업무적합성 검사

            ;교육사항
            table := this.cUIA.FindElement({ AutomationId: "tb7" })
            n := 1
            row := table.Findall({ LocalizedType: "행" })[n + 1]
        }
        else {
            ;인수인계
            table := this.cUIA.WaitElement({ AutomationId: "tb1" }, 2000)
            Sleep 250
            try
                table.FindAll({ AutomationId: "I_INGYEP" })[2].value := this.workers.firstworker
            catch {
                this.cUIA.WaitElement({ AutomationId: "btnAdd1" }, 2000).Invoke()
                Sleep 250
                table.FindAll({ AutomationId: "I_INGYEP" })[2].value := this.workers.firstworker
            }
            table.FindAll({ AutomationId: "I_INSUP" })[2].value := this.workers.handover
            table.FindAll({ AutomationId: "I_ITEXT" })[2].value := "특이사항 없음."

            ;교육사항
            this.cUIA.FindElement({ AutomationId: "btnAdd7" }).Invoke()
            sleep 250
            rows := this.cUIA.FindElement({ AutomationId: "tb7" }).Findall({ LocalizedType: "행" })
            n := rows.Length - 1
            row := rows[n + 1]
        }

        ;작업 전 교육 내용
        row.WaitElement({ AutomationId: "I_EDUCASE", mm: 1 }, 50).value := "작업 전 안전교육"
        row.WaitElement({ AutomationId: "I_EDUDAY", mm: 1 }, 50).value := FormatTime(, "yyyy-MM-dd")
        row.WaitElement({ AutomationId: "I_EDUSARAM", mm: 1 }, 50).value := this.workers.tutor
        row.WaitElement({ AutomationId: "I_STTIME", mm: 1 }, 50).value := InStr(this.workers.schedule, "주간") ?
            "09:00" :
            "18:00"
        row.WaitElement([{ AutomationId: "I_EDTIME", mm: 1 }, { AutomationId: "I_ENDTIME", mm: 1 }], 50).value := InStr(
            this.workers.schedule, "주간") ?
            "09:10" :
            "18:10"
        row.WaitElement({ AutomationId: "I_DAESANG", mm: 1 }, 50).value := this.workers.TotalCnt
        row.WaitElement({ AutomationId: "I_ATTQTY", mm: 1 }, 50).value := this.workers.attendCnt
        if this.workers.absentCnt
            row.WaitElement({ AutomationId: "I_NTATTQTY", mm: 1 }, 50).value := this.workers.absentCnt
        row.WaitElement({ AutomationId: "I_SAYU", mm: 1 }, 50).value := this.workers.getReasons()
        row.WaitElement({ AutomationId: "I_NAEYONG", mm: 1 }, 50).value := safeContents[A_WDay]

        ;기타사항 음주
        chkDrink := this.data.Has("chkDrinkDetect") ? this.data["chkDrinkDetect"] : false
        if InStr(this.workers.schedule, "주간") && chkDrink {
            ;기타사항 한줄 추가
            this.cUIA.FindElement({ AutomationId: "btnAdd8" }).Invoke()
            Sleep 250
            ;교정 여부
            chkCalib := this.data.Has("chkDrinkDetectorCalibration") ? this.data["chkDrinkDetectorCalibration"] :
                false
            drinkDetector := chkCalib ? "음주육안검사" : "음주감지기"

            table := this.cUIA.WaitElement({ AutomationId: "tb7" })

            ;첫째줄
            table.FindAll({ AutomationId: "I_NAEYONG" })[1].value := "음주측정 : 검사시행자(" this.workers.tutor "), 일시(" A_YYYY "." A_MM "." A_DD " 09:10), 대상자(" this
            .workers.drinkAttendees ") - " drinkDetector ", 정상"
            ;둘째줄
            table.FindAll({ AutomationId: "I_NAEYONG" })[2].value := "음주측정 : 검사시행자(" this.workers.secondChecker "), 일시(" A_YYYY "." A_MM "." A_DD " 09:10), 대상자(" this
            .workers.tutor ") - " drinkDetector ", 정상"
        }

        ;저장
        this.cUIA.FindElement({ AutomationId: "btnSave" }).Invoke()
        this.cUIA.WaitElement({ LocalizedType: "단추", Name: "확인" }, 3000).Invoke()
        this.cUIA.WaitElement({ LocalizedType: "단추", Name: "확인" }, 3000).Invoke()
    }

    SafeManage() {
        ; 데이터 소스: this.data["safetyList"] (가정)
        ; 구조: [{content: "", start: "", end: "", confirm: []}, ...]
        safetyList := this.data.Has("safetyList") ? this.data["safetyList"] : []

        if (safetyList.Length == 0)
            return

        menuClick(this.cUIA, "안전관리")							;메뉴진입

        ; 안전관리 탭 ID가 tb1, tb2, tb3, tb4임.
        ; loop 대신 safetyList를 순회하며 입력.

        for idx, item in safetyList {
            startStr := item.Has("start") ? item["start"] : ""
            endStr := item.Has("end") ? item["end"] : ""
            content := item.Has("content") ? item["content"] : ""
            confirm := item.Has("confirm") ? item["confirm"] : "allOK"

            if (startStr == "" || endStr == "")
                continue

            startHour := Number(SubStr(startStr, 9, 2))
            if startHour < 8
                count := 4				;새벽
            else if startHour < 12
                count := 1				;오전
            else if startHour < 18
                count := 2				;오후
            else
                count := 3				;저녁

            ; 해당 시간대 테이블 찾기
            table := this.cUIA.WaitElement({ AutomationId: "tb" count }, 2000)

            ;작업내용
            table.FindElement({ AutomationId: "I_ORTEXT" }).value := content
            ;시작일
            table.FindElement({ AutomationId: "I_STARTDAY" count }).value := SubStr(startStr, 1, 4) "-" SubStr(
                startStr,
                5, 2) "-" SubStr(startStr, 7, 2)
            ;시작시간
            table.FindElement({ AutomationId: "I_STARTTIME" count }).value := SubStr(startStr, 9, 2) ":" SubStr(
                startStr, 11, 2)
            ;종료일
            table.FindElement({ AutomationId: "I_ENDDAY" count }).value := SubStr(endStr, 1, 4) "-" SubStr(endStr,
                5, 2
            ) "-" SubStr(endStr, 7, 2)
            ;종료시간
            table.FindElement({ AutomationId: "I_ENDTIME" count }).value := SubStr(endStr, 9, 2) ":" SubStr(endStr,
                11,
                2)
            ;책임자
            table.FindElement({ AutomationId: "I_OFFICER" }).value := this.workers.firstWorker
            ;작업자
            table.FindElement({ AutomationId: "I_WORKER" }).value := this.workers.workers
            ;불러오기
            table.FindElement({ AutomationId: "btnOpen" count }).Invoke()
            Sleep 250

            ;일상점검 시 (새벽시간대 제외)
            if count != 4 {
                rows := this.cUIA.WaitElement({ AutomationId: "tbs" count }, 2000).Findall({ LocalizedType: "행" })
                ; (6,9,10,11행은 N/A)
                NA := [6, 9, 10, 11]

                loop 4 {

                    rows[NA[A_Index] + 1].WaitElement({ AutomationId: "I_DAYCHECK" }, 100).setFocus()
                    this.cUIA.send("{up 2}")
                }
            }
        }

        this.cUIA.FindElement({ AutomationId: "btnSave" }).Invoke()
        this.cUIA.WaitElement({ LocalizedType: "단추", Name: "확인" }, 3000).Invoke()

    }

    DrivingCheck() {
        chkTime := this.data.Has("drivingTime") ? this.data["drivingTime"] : "18:10"
        chker := this.data.Has("drivingChecker") ? this.data["drivingChecker"] : ""

        ; 운전자 정보는 workers 객체에서 가져옴 (생성자에서 drivers 리스트로 초기화됨)
        drivers := this.workers.drivers

        thisID := WinExist("A")
        menuClick(this.cUIA, "운전적합성 점검표")			;메뉴진입

        ; 로컬 함수 정의 (thisID 사용 불가하므로 직접 호출)
        ; elementFuncScript는 외부 함수이므로 그대로 사용 가능하나 thisID 참조 주의
        ; 여기선 JS 호출이 필요함.

        this.cUIA.WaitElement({ Type: "Image", Name: "피검사자" }).Invoke()
        lensInput(thisID, , drivers[1])

        ;검사자 지정여부
        this.cUIA.WaitElement({ Type: "Image", Name: "검사자" }).Invoke()
        if chker or drivers.Length == 1
            lensInput(thisID, , chker)				;검사자 입력
        else
            lensInput(thisID, , drivers[2])			;운전자2 입력

        if drivers.Length == 2 {					;운전자2가 있을 때
            this.cUIA.WaitElement({ AutomationId: "btnAdd1" }).Invoke()
            sleep 250

            this.cUIA.Findall({ Type: "Image", Name: "피검사자" })[2].Invoke()
            lensInput(thisID, , drivers[2])				;운전자2 입력

            this.cUIA.Findall({ Type: "Image", Name: "검사자" })[2].Invoke()
            if chker
                lensInput(thisID, , chker)				;검사자 입력
            else
                lensInput(thisID, , drivers[1])			;운전자1입력
        }

        loop drivers.Length {
            n := A_Index
            ;날짜
            this.cUIA.WaitElement({ AutomationId: "I_ZCHECK_DAY" n }, 100).value := FormatTime(, "yyyy-MM-dd")
            ;시간
            this.cUIA.WaitElement({ AutomationId: "I_ZCHECK_TIME" n }, 100).value := chkTime

            ;적합 x8
            loop 7 {
                this.cUIA.WaitElement({ AutomationId: "I_ZCHECK_" A_Index "00" n }, 100).setFocus()
                this.cUIA.send("{PgDn}")
            }
            this.cUIA.WaitElement({ AutomationId: "I_ZSUIT" n }, 100).setFocus()
            this.cUIA.send("{PgDn}")

        }
        Sleep 250

        ;저장
        this.cUIA.FindElement({ AutomationId: "btnSave" }).Invoke()
        this.cUIA.WaitElement({ LocalizedType: "단추", Name: "확인" }, 3000).Invoke()
    }

}
