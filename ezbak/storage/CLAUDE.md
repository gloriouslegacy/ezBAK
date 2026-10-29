# ezbak/storage/
역할: 백업 저장 위치·방식(V2)이 늘거나 바뀔 때 수정된다(로컬·외장·UNC, NAS 연결, 이후 클라우드·압축)
공개 진입점(하위 모듈 nas 는 `__all__` 에 있어 `from ezbak.storage import nas` 허용, ADR 0002):
- LocalFolderStore(dest, clock=datetime.now) — core.ports.BackupStore 구현(new_backup_path·new_log_path·list_backups·delete_backup(name, on_log)·free_space)
- nas.NasSpec(unc, drive=None, user="", password=None, persistent=True)(password 는 repr=False), nas.NetUseResult(rc, out, err, message), nas.build_connect_cmd(spec), nas.disconnect_target(text) -> (argv, label)|None, nas.run_net_use(argv) -> NetUseResult, nas.free_drive_letter(used)
의존: 들어오는 것 gui/, automation/ · 나가는 것 core·winsys 의 공개 이름(core.ports·model·rmtree_readonly, winsys.free_space·proc) · 금지: tkinter, subprocess 직접, targets/·automation/ import
| 파일 | 역할 |
|---|---|
| local.py | 새 백업 폴더·로그 경로(<user>_backup_<YYYY-MM-DD_HHMMSS>, 경로만 계산하고 만들지 않음), 백업 목록(isdir, 소문자 접두 비교, 이름 정렬 — GUI·Headless 같음), 삭제(rmtree_readonly), 여유 공간 |
| nas.py | net use 명령 조립(연결·해제 대상 해석), run_net_use(cp949·errors=ignore, 예외면 같은 명령을 인코딩 없이 한 번 더, 두 번째 예외는 올림), 빈 드라이브 문자 고르기(Z..D, C, B, 'Z' 현행 순서). 사용 중 문자 조회는 winsys.sysinfo.used_drive_letters |
테스트: python -m pytest tests -q -k storage
규칙: 백업 폴더·로그 이름 규칙과 보관 정리의 이름 정렬·접두 비교는 C3 고정. NAS 암호는 argv 로만 넘기고 repr·message·예외 문구·진단 로그에 넣지 않는다(진단은 예외 type·rc 만)
NAS 대화상자 UI·큐 문구·예외 처리는 gui/dialogs/nas.py(NasDialogsMixin)에 있다(P4 결정 D3). run_net_use 가 예외를 올리면 gui 는 큐 로그와 오류 대화상자 하나로 끝낸다(자격 증명 재입력·재실행 없음, P4-fix F1). 이 폴더는 실행과 결과 해석만
폴더 만들기·같은 이름 폴더 삭제는 진입 wrapper 몫(ADR 0002 D2). 폴더·로그 경로는 clock 을 각각 읽는다(현행: 초가 바뀌면 이름이 다를 수 있음)
delete_backup 의 항목별 삭제 실패는 예외가 아니라 on_log 메시지로 간다(rmtree_readonly 동작). 그 밖의 예외는 호출처로 올린다
새 저장 방식은 BackupStore 구현 파일 하나를 추가한다. 엔진(core)을 고치게 되면 경계가 틀린 것이니 보고한다
ADR: docs/adr/0002-패키지-구조.md
