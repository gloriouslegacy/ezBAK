# ezBAK 리팩터링 요구사항

출처: 인터뷰(2026-10-12), 이전 계획서 `ezBAKrefactoringPlan.md`(2026-09-03, 흡수 후 삭제 예정)

## 목적
약 8,200줄짜리 단일 파일 `ezBAK.py` 를 동작·호환성 변화 없이 모듈로 나눈다. 그래서 백업 대상, 저장 방식, 자동화 기능을 추가할 때 한 곳만 고치면 되게 한다.
성공 기준:
- `ezBAK.py` 는 100줄 이하 진입 wrapper 가 된다.
- GUI·Headless·스케줄이 같은 백업 엔진을 쓴다.
- 기본 필터는 한 곳에서만 정의한다.
- 자동 테스트가 통과한다.
- PyInstaller 빌드가 성공한다.

## 사용 시나리오
- S1 GUI 에서 사용자를 고르고 데이터를 백업한다. 숨김·시스템 파일 포함 여부, 포함·제외 필터, 브라우저 북마크, 보관 개수 정리가 포함된다.
- S2 GUI 에서 백업 폴더를 골라 사용자 데이터를 복원한다.
- S3 Windows 예약 작업이 `ezBAK.exe` 를 인자와 함께 Headless 로 실행한다. 스케줄 config JSON 대로 백업하고 오래된 백업·로그를 정리한다.
- S4 드라이버를 백업·복원한다(pnputil). 설치 앱 목록을 내보낸다(winget). 브라우저 프로필 전체를 백업한다.
- S5 NAS 를 연결·해제한다. 데이터 복사, 공간 확인, 장치 관리자·탐색기 열기를 한다.
- S6 개발자가 태그(`v*`)를 push 하면 GitHub Actions 가 onefile exe 를 빌드해 릴리스한다.

## 품질 속성
- Q1 동작 동일성: 리팩터링 전후로 GUI·Headless 백업 결과가 같아야 한다(폴더 구조, 파일 수, 로그 형식). 단, Phase 0 에서 의도해 고친 부분은 예외다.
- Q2 회귀 검증은 자동 테스트로만 한다(수동 GUI 확인 없음).
  - 특성 테스트: 필터·is_hidden, 임시 폴더를 쓰는 Headless 백업
  - GUI 스모크: 화면을 띄우지 않고 App 생성·위젯 구성·queue 처리를 확인한다.
  - 외부 명령(schtasks·pnputil·net use·winget·UAC): subprocess 를 모킹해 만들어진 명령줄만 확인한다. 실제 등록·설치는 하지 않는다.
- Q3 오류를 숨기지 않는다: `except: pass` 와 `DEBUG:` print 를 로깅 체계로 바꾼다.
- Q4 성능 수치 목표는 없다. 현재보다 느려지지 않으면 된다.

## 제약
- C1 Windows 10/11, Python 3.11, tkinter. pywin32·psutil 은 선택이다(없으면 fallback 동작).
- C2 배포는 PyInstaller `--onefile --windowed --uac-admin`, exe 이름 `ezBAK` 이다. `.github/workflows/release.yml` 이 기준이며, 로컬 빌드 스크립트·spec 을 여기에 맞춘다(지금은 아이콘 데이터와 버전 파일이 다르다).
- C3 호환성은 전부 유지한다.
  - exe 이름, CLI 인자, settings·schedule config JSON 형식을 바꾸지 않는다.
  - 이미 등록된 예약 작업은 수정 없이 계속 동작해야 한다.
- C4 하지 않을 것: 프레임워크 교체, UI 리디자인, 백업 알고리즘 변경(압축·증분), 한 번에 전체 분리
- C5 1인 개발이다. `refactor/package-split` 브랜치에서 발주 1건마다 커밋하고, Phase 가 끝나면 main 에 병합한다.
- C6 루트 정리
  - 스크립트(ps1·sh)는 `scripts/`, 테스트는 `tests/`, 문서는 `docs/` 로 옮긴다.
  - `ezBAK.py` 는 루트 진입점으로 남긴다.

## 변경 축
- V1 백업 대상 종류가 늘어난다(브라우저 추가, 앱 설정 등).
- V2 저장 위치·방식이 늘어난다(NAS 외 클라우드·외장, 압축·증분). 이번에 구현하지는 않지만 경계는 둔다.
- V3 자동화·CLI 가 늘어난다(스케줄 옵션, CLI 인자, 원격·일괄 배포).
- 고정: UI 구성, 언어(한/영), 테마. 번역을 JSON 으로 빼지 않는다.

## 미정
- 로그 형식·상관 ID, 줄 수 상한 값, 승인할 자동 검사 → ADR 에서 정한다.
- Phase 0 에서 필터 타입을 `'NAME'` 에서 `'name'` 으로 통일했다. 이전 settings JSON 에 `'NAME'` 이 저장돼 있으면 그대로 읽히는지 확인하고, 필요하면 마이그레이션한다(C3).
