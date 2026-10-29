# P5-fix 계획 (초안 2026-10-13 · 기준 6c56d40) — roadmap "P5 진행 중 고정한 이상 동작" 20건 + 표 밖 발견 4건, 커밋 = 발주 1건
사용자 결정(2026-10-13): roadmap 의 항목은 P6 전에 모두 고친다. 아래 D1~D8 은 "어떻게"만 묻는다(권장안 기준으로 커밋을 짰다).
공통: 결함 수정(Q1 예외) 커밋과 정리(동작·golden 무변경) 커밋을 섞지 않는다. golden 재생성은 그 커밋의 테스트 파일만(`set EZBAK_REGEN_GOLDEN=1 && python -m pytest <파일>`), 두 번 생성해 같은지·키 단위로 비교하고 바뀐 키를 커밋 메시지에. 재생성물은 5개 셈에서 빼지만 병렬 판단에서는 겹침으로 센다.
gui 제약: import 별칭(threading·datetime·os·filedialog·Win11Dialog 등) 이름 유지(facade 가 이름으로 patch). 새 helper 는 App 메서드가 아니라 모듈 함수, App 인스턴스 속성을 새로 만들지 않는다(app_unchanged). 진단은 모듈 `_log = logging.getLogger(__name__)`(actions_restore 와 같은 이름)에 `event=… error=%s` 형식. 새 번역 키는 en·ko 둘 다 + gui_snapshot `translations` 재생성. 실제 schtasks·드라이버·UAC·네트워크·사용자 폴더 금지(block_subprocess·isolated_env).
C3 표시: ● C3 항목(키 이름·큐 task·작업 로그 문구·JSON) 변경 · △ 번역 값만(키 이름 그대로)·키 추가 · – 없음

## 1. 항목 표 (위치는 6c56d40 에서 확인)
| id | 현재 동작 | 위치 | 고칠 것(최소) | C3 | golden(파일: 키) | 결정 |
|---|---|---|---|---|---|---|
| h-a | `<Control-1>`~`<Control-5>` 가 Ctrl+마우스 버튼 1~5 로 묶임(Ctrl+6~9 만 키). Ctrl+클릭하면 사용자 1~3 번이 선택됨 | gui/shortcuts.py `KeyboardShortcuts.setup_shortcuts` 76-77 | `f'<Control-Key-{i}>'` | △(D1=A 면 도움말 값) | gui_snapshot: bindings.keys(`<Control-Button-1..5>`→`<Control-Key-1..5>`), D1=A 면 translations.tables.en·ko `keyboard_shortcuts_content` | D1 |
| h-b | `<Space>` 는 Tk 에 없는 keysym 이라 TclError → except 로 빠져 그 뒤 Ctrl+E·Ctrl+C·F5·F1 이 안 묶임(Space 도 안 묶임) | gui/dialogs/sources.py `_setup_source_shortcuts` 72 | `'<space>'` | – | gui_dialogs: sources_tree_dark_en·sources_tree_light_ko 의 bindings·snapshot | — |
| h-c | `<Alt-1/2/3>` 가 Alt+마우스 버튼(Alt+클릭이 주기를 바꿈) | gui/dialogs/schedule.py `_setup_schedule_shortcuts` 129-131 | `'<Alt-Key-1>'`… | – | gui_dialogs: schedule_tree_dark_en·schedule_tree_light_ko 의 bindings·snapshot | — |
| h-d | 스케줄 생성 예외 대화상자 "Failed to create schedule: {0}: <오류>" | dialogs/schedule.py `on_create` except 292 | `get('failed_to_create_schedule').format(create_error)` | – | gui_dialogs: schedule_create_bad_retention.dialogs | — |
| h-e | 대상 폴더가 없을 때 "Existing data for '{0}' will be overwritten…?:\n<dest>\n\nContinue anyway?"(뜻도 틀림) | dialogs/schedule.py `on_create` 261 | D2=A: 새 키 `dest_folder_missing`("The destination folder does not exist" / "대상 폴더가 없습니다") 로 `f"{새 키}:\n{dest}\n\n{continue_anyway}"` | △ 키 추가(기존 키 그대로) | gui_dialogs: schedule_create_missing_dest_no·_yes.dialogs / gui_snapshot: translations | D2 |
| h-f | en 사전 중복 키 25개 중 값이 같은 19개(no_log_content·task_creation_failed·…·system_mismatch_warning) | gui/strings_en.py 134-164 의 뒤 줄 | 뒤 줄 19개 삭제(dict 는 첫 삽입 위치·마지막 값 → 순서·값 그대로) | – | 없음(translations 무변경 확인) | — |
| h-g | 값이 다른 6개는 뒤의 `{0}` 없는 값이 쓰이고, 호출처 4개는 f-string 이어 ko 에 `{0}` 가 남음: error_saving_log(en 은 오류 내용 빠짐), failed_to_delete_task("…task::\n<e>", ko "{0}:\n<e>"), task_deleted("deleted.: 'x'"), source_selection_failed·filter_manager_error("…::"), invalid_drive_letter("Invalid drive letter.") | strings_en.py 135·142·144-146·156, 호출처 actions_schedule.py 222·240, actions_copy.py 132, app.py 310, dialogs/schedule.py 218 | 뒤 6줄 삭제(첫 값 유지, D3) + 호출처 `.format(e)`·`.format(task_name)`. app_log 183·nas 135/196 은 이미 맞음 | △ en 값 6개 | gui_snapshot: translations.tables.en / schedule_cases: create_then_delete·delete_no_appdata·delete_ok_config_sanitized·delete_ok_no_config·delete_fail_rc1·delete_fail_rc1_no_output .dialogs / gui_cases: space_sources_dialog_error-pywin32·-nopywin32 .dialogs / gui_dialogs: nas_connect_input_checks.steps / tool_net: nas_connect_bad_drive·nas_disconnect_bad_drive .dialogs | D3 |
| h-h | NAS 해제 대상이 비면 showerror("Error","Error") | dialogs/nas.py `open_nas_disconnect_dialog.on_ok` 316 | 본문을 기존 키 `input_target` 로 | – | gui_dialogs: nas_disconnect_input_checks.steps | — |
| h-i | 보관 개수·월 날짜가 숫자가 아니면 `action="create"` 를 먼저 넣고 int() 에서 예외 → 오류 뒤 Alt+F4(destroy 직행)로 닫으면 "No result data found" 분기 | dialogs/schedule.py `on_create` 272-284 | result dict 를 지역 변수로 먼저 만들고 성공한 뒤 `action`·`result` 대입 | – | gui_dialogs: schedule_create_bad_retention.action("create"→null) | — |
| h-j | 드라이브 문자 상자 활성 상태가 `<Map>` 때만 갱신 → 체크 해제해도 입력 가능 | dialogs/nas.py `_connect_fields.toggle_drive` 87-99 | `map_var.trace_add('write', lambda *_: toggle_drive())` | – | 없음(새 테스트: 체크 해제 → disabled, 다시 → normal) | — |
| h-k | 원본 선택 결과가 set 순서, 고른 루트는 `/`(askdirectory) + 하위는 `\` 로 섞임 | dialogs/sources.py `_on_ok` 299-320·`_choose_root` 285 | 루트를 `os.path.normpath`, `_on_ok` 는 트리 순서(get_children 깊이 우선)로 모음 | – | gui_dialogs: sources_select_and_ok.steps·selected_sorted(경로 구분자). 순서는 golden 이 정렬하므로 새 테스트로 | — |
| h-l | 복사 계산 총량 < 복사 바이트(132 vs 145): 파일 원본을 `os.walk` 해 "Walk error [WinError 267]" 후 0 으로 셈 | gui/actions_copy.py `_copy_space_ok` 207-211 | 파일 원본은 `run_space_calculation` 처럼 `os.path.isfile` 이면 `getsize`(is_hidden 이면 제외), 폴더만 `compute_total_bytes_safe` | – | gui_cases: copy_tree·copy_include_hidden_system·copy_insufficient_space·copy_disk_usage_fail·copy_missing_source(-pywin32·-nopywin32) .logs·.queue | — |
| h-m | pywin32 없으면 크기 계산(`is_hidden_safe`, ctypes 대체·AppData 이름 없음)이 숨김 파일도 셈(212 vs 145). 복사·공간 확인은 `is_hidden` | gui/actions_backup.py `compute_total_bytes_safe` 219 | 판정을 `self.is_hidden(p)` 로(복사와 같은 판정). `is_hidden_safe` 는 그대로(test_filters golden 'safe' 열) | – | gui_cases: 위 copy_* -nopywin32(.logs·.queue) | D4 |
| h-n | 복사 finally 가 작업 스레드에서 `progress_bar['maximum']` 를 읽음 → 큐의 update_progress_max 가 아직 처리 전이라 기본값 100 을 보냄 | actions_copy.py `run_copy_thread` finally 293-297 | D5=A: 그 블록 삭제(성공 경로는 try 안의 `update_progress, bytes_copied` 가 h-l 뒤 총량과 같아 꽉 참) | – | gui_cases: 위 copy_* .queue 의 끝 `["update_progress",100]` 삭제 | D5 |
| h-o | `_on_open` 숨김 판정 예외를 pass(항목은 보임) | dialogs/sources_tree.py 94-95 | 동작 그대로, `_log.debug("event=sources_hidden_check_failed path=%s", full, exc_info=True)` | – | 없음 | — |
| h-p | `_refresh_tree` 의 checked_paths 를 모으고 안 씀(F5 는 h-b 뒤 처음 동작: 체크·추가 루트가 사라짐) | dialogs/sources.py 184-188 | D6=A: 미사용 목록 삭제(정리) | – | 없음 | D6 |
| h-q | 지역 `from datetime import datetime`(mod.datetime 패치가 안 닿음. fixed_clock 은 datetime 모듈을 바꿔 결과는 같음) | dialogs/schedule.py `_schedule_time_str` 232, actions_copy.py `_copy_to_new_folder` 236 | 지역 import 삭제, schedule.py 는 맨 위 `from datetime import datetime`(별칭 이름 그대로) | – | 없음 | — |
| h-r | DEBUG print: 복사 흐름 14곳(gui_cases stdout 에 기록), on_create 4곳 | actions_copy.py `copy_files`·`check_space`·선택 helper, dialogs/schedule.py `on_create` | D7=A: `_log.debug` 로(형식 ADR 0003 §2). 나머지 ~120곳은 P6 검사 d | – | gui_cases: copy_* 16키 .stdout(빈 목록이 됨) | D7 |
| h-s | open_log_file 의 맨 `except:` 2곳(retention 값 읽기, 대체값 '2'·'30') | gui/app_log.py 106·110 | `except Exception:`(대체값 그대로) | – | 없음 | — |
| h-t | 위젯 꾸미기가 아닌 `except Exception: pass` 30곳(ADR 0003 §2 위반): 작업 로그 쓰기·머리말·닫기(app_log 90·126·140, actions_backup 173, actions_copy 290, actions_restore 88, actions_tools 43·84), 설정 읽기·저장(app 99·305, i18n 125, theme 117, layout 178), 결과 폴더 열기(bridge 174), 대화상자 표시 실패(app 312, actions_tools 382, actions_schedule 215·224·242, nas 160·165·173·180·186·206·213·219, schedule 57·112·220) | 왼쪽 | `pass` → `_log.debug("event=<무엇> failed", exc_info=True)`. 동작·주석 의미 그대로. 아이콘·grab·focus·진행 막대·버튼·중앙 정렬은 그대로 | – | 없음 | — |
표 밖 발견(roadmap 에 없음, D8): h-u 복사 끝 상태가 늘 "Operation finished"(`_Copy_completed` 를 세우는 곳이 없어 "Copy complete!" 를 덮음, actions_copy 301 · golden copy_tree 등 .queue 끝) · h-v shortcuts.py 맨 `except:` 3곳(104 아이콘·188 종료 때 설정 저장·202 상태 표시줄) · h-w 백업·복원·도구 finally 도 작업 스레드에서 `progress_bar['maximum']` 를 읽음(h-n 과 같은 경쟁, actions_backup 177·actions_restore 97·actions_tools 88·297) · h-x 메인 창 Ctrl+C(copy_files)가 루트에 묶여 로그 글상자에서 복사해도 복사 기능이 뜸

## 2. 결정 (사용자)
- **확정 2026-10-13: D1~D7 모두 A, D8 은 h-u A·h-v A·h-w B·h-x B (전부 권장안)**
- D1 도움말 `keyboard_shortcuts_content` 의 "Ctrl + 1~9 (Not Working)"(h-a 뒤 거짓이 됨): A(권장) 그 줄의 "(Not Working)"/"(작동 안 함)" 만 지움(상자 폭 유지, F5·Escape 표시는 그대로) / B 그대로 — 고친 기능을 "안 됨"으로 안내하게 된다
- D2 대상 폴더 없음 문구(h-e): A(권장) 새 키 `dest_folder_missing` + 기존 `continue_anyway` / B `overwrite_warning.format(user)` / C `.format(dest)` — B·C 는 `{0}` 만 없애고 "덮어쓴다"는 틀린 뜻과 질문 두 개가 남는다
- D3 en 값이 다른 6키: 지금 쓰이는 값은 모두 뒤 값(`{0}` 없음). A(권장) 6개 모두 첫 값(`{0}` 판, ko 와 같은 뜻) 유지 + 호출처 `.format` / B invalid_drive_letter 만 뒤 값 "Invalid drive letter." 유지(golden 3키 덜 바뀜, ko·`valid_drive_letter` 와 다름) — 나머지 5개는 뒤 값이면 `.format` 결과에 내용이 빠져 A 뿐이다
- D4 크기 판정(h-m): A(권장) `compute_total_bytes_safe` 가 `is_hidden` 을 씀 — 복사·공간 확인과 같은 판정, test_filters golden 무변경 / B `is_hidden_safe` 에 ctypes 대체 추가 — test_filters 'safe' 열이 바뀌고 AppData 이름 차이는 남는다
- D5 복사 마지막 진행 채우기(h-n): A(권장) finally 블록 삭제 — 성공은 이미 꽉 차고, 실패·공간 부족은 멈춘 자리에 남음 / B 총량을 반환값으로 들고 와 finally 에서 총량으로 채움(`_copy_space_ok`·`_copy_to_new_folder` 반환값 변경) — B 는 실패도 꽉 찬 막대로 보인다
- D6 checked_paths(h-p): A(권장) 미사용 목록 삭제, F5 는 지금 코드대로 체크·추가 루트를 지움 / B 새로고침 뒤 추가 루트를 다시 넣고 그 루트 체크를 복원(동작 변경·새 테스트·sources.py +15줄 안팎)
- D7 DEBUG print(h-r): A(권장, "P6 전에 모두" 결정에 맞춤) 복사 흐름·on_create 만 지금 `_log.debug`, 나머지는 P6 검사 d / B 이 항목도 P6 d 로 일괄 — B 는 gui_cases stdout golden 을 P6 에서 한 번만 바꾼다
- D8 표 밖 발견: h-u A(권장) 이번에 포함(별도 Q1 커밋, 성공이면 "Copy complete!") / h-v A(권장) h-s 커밋에 포함(같은 규칙, 아이콘·상태 표시줄은 주석 달고 `except Exception: pass`, 종료 저장은 `_log.debug`) / h-w B(권장) P6 로(golden 4파일, 진단 전환과 함께) / h-x B(권장) 그대로(단축키 체계 변경 = UI 동작)

## 3. 커밋 (권장안 기준. ∥ = 병렬 가능)
| # | 커밋 | 항목 | Files(≤5, golden 제외) | Q1 | 새 테스트 |
|---|---|---|---|---|---|
| 1 | fix-ctrlnum | h-a | gui/shortcuts.py, gui/strings_en.py, gui/strings_ko.py | 예: gui_snapshot bindings.keys, translations(en·ko keyboard_shortcuts_content) | golden 으로 |
| 2 | fix-copytotal | h-l, h-m | gui/actions_copy.py, gui/actions_backup.py, tests/characterization/test_gui_cases.py | 예: gui_cases copy_tree·include_hidden_system·insufficient_space·disk_usage_fail·missing_source ×2 .logs·.queue | copy 성공 경우 update_progress_max == 마지막 update_progress 단언. insufficient_space 가 여전히 부족인지 확인 |
| 3 | fix-dlgkeys | h-b, h-c | gui/dialogs/sources.py, gui/dialogs/schedule.py | 예: gui_dialogs sources_tree_*·schedule_tree_* .bindings·.snapshot | golden 으로 |
| 4 | fix-copyprogress | h-n | gui/actions_copy.py | 예: gui_cases 위 copy_* .queue 끝 update_progress 100 | — |
| 5 | clean-endup | h-f | gui/strings_en.py | 아니오(translations 무변경) | — |
| 6 | fix-fmt | h-g | gui/strings_en.py, gui/actions_schedule.py, gui/actions_copy.py, gui/app.py, gui/dialogs/schedule.py | 예: 표 h-g 의 5파일 키 | — (ko 는 golden 없음, 커밋 메시지에) |
| 7 | fix-schedmsg | h-d, h-e | gui/dialogs/schedule.py, gui/strings_en.py, gui/strings_ko.py | 예: gui_dialogs schedule_create_bad_retention·missing_dest_no·_yes .dialogs, gui_snapshot translations(키 추가) | — |
| 8 | fix-schedaction | h-i | gui/dialogs/schedule.py | 예: gui_dialogs schedule_create_bad_retention.action | — |
| 9 | fix-debugprint | h-r | gui/actions_copy.py, gui/dialogs/schedule.py | 예: gui_cases copy_* .stdout | — |
| 10 | fix-copydone | h-u | gui/actions_copy.py | 예: gui_cases 복사가 끝나는 copy_* .queue 끝 update_status | 반환값으로 완료 여부(App 속성 아님) |
| 11 | fix-nasempty | h-h | gui/dialogs/nas.py | 예: gui_dialogs nas_disconnect_input_checks.steps | — |
| 12 | fix-nasdrive | h-j | gui/dialogs/nas.py, tests/characterization/test_gui_dialogs.py | 아니오 | map 체크 토글 → 상자 state |
| 13 | fix-sources | h-k | gui/dialogs/sources.py, tests/characterization/test_gui_dialogs.py | 예: gui_dialogs sources_select_and_ok.steps·selected_sorted | 결과 순서 = 트리 순서, 구분자 `\` 만 |
| 14 | clean-local | h-q, h-p | gui/dialogs/schedule.py, gui/actions_copy.py, gui/dialogs/sources.py | 아니오 | — |
| 15 | clean-bareexcept | h-s, h-v | gui/app_log.py, gui/shortcuts.py | 아니오 | — |
| 16 | clean-swallow-app | h-t(설정·폴더 열기) | gui/app.py, gui/i18n.py, gui/theme.py, gui/layout.py, gui/bridge.py | 아니오 | — |
| 17 | clean-swallow-jobs | h-t(작업 로그) | gui/app_log.py, gui/actions_backup.py, gui/actions_copy.py, gui/actions_restore.py, gui/actions_tools.py | 아니오 | — |
| 18 | clean-swallow-dlg | h-t(대화상자), h-o | gui/actions_schedule.py, gui/dialogs/nas.py, gui/dialogs/schedule.py, gui/dialogs/sources_tree.py | 아니오 | — |
순서(같은 파일·같은 golden JSON 은 순차. golden 을 재생성하는 커밋 옆에는 그 golden 이 실행하는 파일을 고치는 커밋을 두지 않는다 — 작업 트리 공유. strings_en/ko 를 고치는 1·5·6·7 은 모든 golden 이 읽으므로 단독):
- 1단계 Q1: ① 2 ∥ 3 (gui_cases 는 원본 대화상자를 대역으로, gui_dialogs 는 복사 코드를 안 부름) → ② 4 ∥ 11 → ③ 1 → ④ 5 → ⑤ 6 → ⑥ 7 → ⑦ 8 → ⑧ 9 → ⑨ 10 ∥ 12 (12 는 golden 재생성 없음, 파일 겹침 없음) → ⑩ 13
- 추가(2026-10-13): #12 뒤 test_gui_dialogs.py 가 500줄 → #13 전에 #12b split-dialogs-test(NAS 절을 새 파일로, golden 무변경)를 ⑨ 에서 10 과 병렬로
- 2단계 정리(golden 무변경 확인만, 1단계가 다 끝난 뒤): 14 ∥ 15 ∥ 16 → 17 ∥ 18. 각 묶음 뒤 runner 로 전체 pytest
- D 가 권장안과 다르면: D1=B 면 1 에서 strings 2파일 빠짐 / D5=B 면 4 에 반환값 변경 / D6=B 면 14 에서 h-p 를 빼 Q1 커밋으로(sources.py, test_gui_dialogs.py, 13 뒤) / D7=B 면 9 삭제 / D8 h-u=B 면 10 삭제, h-v=B 면 15 는 app_log.py 만

## 4. 완료 기준
- `python -m pytest tests -q -W error` 전부 통과, golden 변경은 위 표의 키뿐(Q1 커밋마다 키 단위 비교), 정리 커밋 14~18 은 golden 무변경
- folder-check 위반 0, 아키텍처 검사 a~c 통과, 파일 500줄 이하(strings_en 287 → 약 262, schedule.py 346·actions_copy 304 는 증가 10줄 안쪽)
- 끝나면 메인: roadmap 에 P5-fix 행(커밋 해시·D 결정), 결함 문단의 처리 결과 / gui CLAUDE.md 의 "en 중복 키 25개" 줄과 "옮긴 pass 는 주석만" 줄 갱신 / 표 밖 h-w·h-x 를 roadmap "함께 처리할 것" 으로
