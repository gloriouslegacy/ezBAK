# 구조 색인
상태: 확정(ADR 0002 승인 2026-10-12)      갱신: 2026-10-27
리팩터링 P1~P5 로 아래 구조가 됐다(루트 `ezBAK.py` 69줄 + `ezbak/` 6개 모듈). 파일별 역할은 각 모듈의 `CLAUDE.md` 를 본다.

| 모듈 | 역할(무엇이 바뀔 때 수정되나) | 의존(나가는 것) |
|---|---|---|
| `ezBAK.py` (루트) | 진입 wrapper·PyInstaller entry. Headless 판정·UAC·GUI 시작 순서가 바뀔 때 | automation, gui, winsys |
| `ezbak/core/` | 백업·복원 절차, 필터·숨김 규칙, 보관 정리, 작업 로그 문구, 도메인 타입·포트, 진단 로그 설정(`diag`, `%APPDATA%\ezBAK\logs\ezbak.log`) | winsys(fileattrs·paths 만) |
| `ezbak/targets/` | 백업 대상 종류(V1): 브라우저 북마크·프로필, 드라이버(pnputil), 앱 목록(winget) | core, winsys |
| `ezbak/storage/` | 저장 위치·방식(V2): 로컬 폴더 저장소, NAS 연결·해제(net use) | core, winsys |
| `ezbak/automation/` | 자동화·CLI(V3): 인자, Headless 실행, 스케줄 JSON, 예약 작업 등록(schtasks) | targets, storage, core, winsys |
| `ezbak/winsys/` | Windows I/O 단일 출구: subprocess, 파일 속성, 경로·환경 변수, UAC, 레지스트리·여유 공간, 셸 열기 | 표준 라이브러리, pywin32·psutil(선택) |
| `ezbak/gui/` | 화면·대화상자·번역·테마·설정 파일(고정 축) | automation, targets, storage, core, winsys |
| `tests/` | 특성·단위·GUI 스모크·아키텍처 테스트 | ezbak 공개 이름, 루트 `ezBAK.py`(특성 테스트) |
| `scripts/` | 빌드·보조 스크립트(앱 코드 아님) | 루트 파일만 |

의존 방향: `gui → automation → {targets, storage} → core → winsys`. 역방향·형제 내부 파일 import 금지. `tkinter` 는 gui 만, `subprocess` 는 `winsys/proc.py` 만.

ADR: [0001 현재 구조](adr/0001-현재-구조.md) · [0002 패키지 구조](adr/0002-패키지-구조.md) · [0003 구현 규칙](adr/0003-구현-규칙.md) · [0004 UI 리디자인](adr/0004-UI-리디자인.md)
