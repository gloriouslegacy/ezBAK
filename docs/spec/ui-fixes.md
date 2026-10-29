# UI 결함 수정 명세 (v0.10.x 패치)

실기기에서 발견한 UI 결함의 수정 기준. 동작 golden·C3 호환은 바꾸지 않는다. 창 구조 golden 이 바뀌면 Q1 예외 커밋으로 재생성한다.

## F1. Backup Options 두 줄 배치 (2026-10-27, 사용자 결정)

- 증상: Backup 화면 Backup Options 줄이 기본 창 폭(1000)에서 `System:` 뒤(Include·Exclude)가 잘린다. 폭 1160 이면 보인다.
- 수정: 옵션을 두 줄로 나눈다.
  - 1줄: Backups to Keep(+ `(0=all)`) | Log Days(+ `(0=off)`)
  - 2줄: Hidden Include/Exclude | System Include/Exclude
- 위젯 종류·변수·명령·문구·순서는 그대로 두고 배치만 바꾼다.
- 범위: 이 줄만. 다른 화면 잘림 점검은 하지 않는다(사용자가 실기기에서 보고).
- 완료 기준: 기본 창 크기(1000×700)·한국어·영어에서 System 의 Exclude 라디오 오른쪽 끝이 옵션 카드 안에 들어온다.
