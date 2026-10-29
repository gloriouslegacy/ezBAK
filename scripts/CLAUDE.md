# scripts/
역할: 개발·운영 보조 스크립트. 앱 코드가 아니며, 보조 작업 방식이 바뀔 때 수정된다
공개 진입점: 레포 루트에서 직접 실행한다(예: `.\scripts\git_push.ps1 "msg"`, `./scripts/git_push.sh "msg"`)
의존: 나가는 것 — 레포 루트 파일(ezBAK.py, icon/, bak_version_info.txt) · 금지 — ezbak/ 패키지 import
| 파일 | 역할 | 레포 경로 사용 |
|---|---|---|
| build_pyinstaller_script.ps1 | build/·dist/·*.spec 를 정리하고 `-ScriptName`(루트 기준 경로)을 onefile exe 로 빌드 | 예 |
| git_push.ps1, git_push.sh | 루트에서 `git add .` 후 commit, origin main 으로 push | 예 |
| editRetentionDay.ps1 | 테스트 보조. F:\ 의 백업 로그 파일 수정일을 20일 전으로 바꿔 로그 보관 정리를 시험한다 | 아니오 |
| process_monitoring.ps1 | 프로세스(이름·PID)가 끝날 때까지 감시하다 SMTP 메일을 보낸다 | 아니오 |
규칙:
- 레포 경로를 쓰는 스크립트는 자기 위치에서 레포 루트를 계산한다
- 빌드 기준은 .github/workflows/release.yml 이다. 로컬 빌드 플래그는 그 빌드 단계와 같다(P6-9). 다른 점은 버전 파일(로컬 bak_version_info.txt, CI 는 태그에서 생성)뿐이고 spec 파일은 두지 않는다(P6 결정 D3·D4)
- 비밀번호·토큰을 파일에 적지 않는다(process_monitoring.ps1 의 SMTP 비밀번호는 지금 자리표시자)
테스트: 없음(PowerShell 파서 검사 ParseFile, `bash -n`)
