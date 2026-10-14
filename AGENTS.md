# 프로젝트 작업 안내

## SAP RFC 연동

SAP RFC 연결·조회 작업을 수행하기 전에 `SAP_RFC_라이브러리_활용가이드.md`를 읽는다.

- 사용자 지정 원칙: PC에 기존 설치된 SAP DLL을 AutoHotkey `DllCall`로 직접 호출한다.
- PowerShell `Add-Type` / C# 컴파일 등으로 임시 DLL을 생성하는 연결 방식은 사용하지 않는다. 초기 테스트의 임시 DLL이 알약 백신에 감지되었다.
- 현재 검증된 라이브러리는 `C:\Windows\SysWOW64\sapnwrfc.dll`이며, 호출 프로세스는 32비트 AutoHotkey v2를 사용한다.
- 배포용 SAP RFC 코드는 반드시 32비트 독립 실행 파일 `SAPRFC.exe`로 분리한다. 64비트 `Main.exe`에서 RFC DLL을 직접 로드하거나 배포 PC의 AutoHotkey 설치본으로 `.ahk`를 실행하지 않는다.
- `SAPRFC.exe`는 릴리스 빌드 시 미리 컴파일하고 `세션BG.exe`와 같이 `Main.exe`의 `FileInstall` 대상에 포함한다. 배포 사용자는 `Main.exe`만 전달받아도 실행에 필요한 `SAPRFC.exe`가 자동 설치되어야 한다.
- 인증 티켓·세션 원문을 로그, 문서, 저장소에 남기지 않는다. 백신 비활성화·예외 등록을 해결책으로 사용하지 않는다.
- 조회 요청을 변경·확정·저장·커밋 호출로 확대하지 않는다. 다른 업무 RFC는 실제 인터페이스와 권한을 확인하고 검증된 기능과 구분한다.
