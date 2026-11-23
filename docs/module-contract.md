<!-- 칼퇴 저장소 docs/11_모듈_계약_규격.md의 공개 사본(하위 프로그램 개발자용). 원본이 바뀌면 이쪽도 맞춘다. -->
# 하위 프로그램·모듈 계약 규격 (2026-11-22, 판 v0.1.319)

사용자 결정(2026-11-22): 칼퇴 화면이 바뀌어도 **훈수꾼·캐드선수를 패치하지 않아도 잘 돌아야** 한다. 하위 호환이 안 되면 하나 바꿀 때마다 프로그램 3개를 고쳐야 한다. 모듈화하는 이유가 이것이다.

## 1. 원칙
1. **선언은 내용과 "힌트"만 낸다**: 하위 모듈, 켜기/끄기, 사용법, 패치노트, 설정 연결 명령, 그리고(계약 2, 2026-11-22 사용자 요청 "붙는 쪽이 버튼·탭을 조정") **탭·버튼 힌트 `layout`**. **색·글꼴·간격 같은 "모양"은 넣지 않는다.** 모양은 칼퇴가 공통 틀(`ListFrame`·`ToggleListView`·`AppDetailTabs`·`LayoutButtonsPage`)과 디자인 토큰으로 그리므로, 하위 프로그램이 구성(무엇을 어떤 순서로)을 정해도 칼퇴 화면과 일관성이 유지된다. 아무것도 안 적으면 지금처럼 칼퇴 기본 틀.
2. **칼퇴와 모듈은 `commands`로만 이어진다**: 칼퇴가 모듈을 부르는 길은 선언의 `commands`에 적힌 명령뿐이다. 화면·내부 함수에 기대지 않는다. 지금 명령은 `settings`(전용 설정창을 여는 실행 인자) 하나다.
3. **옛 판도 읽는다**: `contract_version`이 없으면 0(PR#54~#59 때의 선언). 칼퇴는 0도 1도 읽는다. 앞으로 판이 올라가도 옛 칸은 지우지 않는다.
4. **새 칸은 늘 선택 항목**: 칸을 더해도 없는 선언은 그대로 유효하다. 필수 칸은 늘리지 않는다.
5. **모르는 칸은 무시**: 칼퇴보다 새 선언이나 화면 모양 칸이 와도 읽기를 거부하지 않고 아는 칸만 읽는다. 무시한 칸은 개발자 설치 때 한 줄로 알려 준다(설치는 막지 않는다).
6. **되돌릴 수 없는 변경만 `spec`을 올린다**: `spec`이 칼퇴가 아는 값보다 높으면 "더 새 칼퇴가 필요"로 거부한다. 계획은 없다. 일반 추가는 `contract_version`도 올리지 않아도 되고, 올리는 건 "이 칸을 쓰는 선언"임을 알리는 표시일 뿐이다.

## 2. 칸 (`<id>.app.json`, `module.json` 공통)
| 칸 | 계약 | 설명 |
|---|---|---|
| `spec` | 0 | 되돌릴 수 없는 변경 표시(지금 0만) |
| `contract_version` | 1 | 선택. 없으면 0. 이 선언이 따르는 계약 판 |
| `modules`, `settings`, `memoryMb`, `startMs` | 0 | 하위 모듈·설정·계기판(docs/10) |
| `guide`, `changes` | 0 | 사용법·패치노트 파일 경로(docs/09) |
| `settingsArgs` | 0 | 설정창 실행 인자(옛 이름, 계속 읽는다) |
| `commands` | 1 | 선택 **객체**. 키는 아래 §3. 모르는 키·모양이면 없는 것으로 본다 |
| `commands.actions` | 2 | 버튼·탭이 부르는 명령(§3). `layout`이 이 명령 id를 가리킨다 |
| `layout` | 2 | 선택 **객체**. 탭·버튼 힌트(§4) |

## 3. commands
```json
"contract_version": 1,
"commands": { "settings": ["--settings"] }
```
- `settings`: 문자열 배열(1~16개, 각 200자 이하). 있으면 소개 화면 오른쪽에 **[설정창 열기 ↗]**, 같은 exe에 이 인자를 붙여 켠다.
- **우선순위**: `commands.settings`가 있으면 그것, 없으면 옛 `settingsArgs`. 둘 다 있어도 오류가 아니다.
- 모양이 틀린 `commands`(배열·문자열·빈 배열·숫자 요소 등)는 **오류가 아니라 없는 것**으로 보고 `settingsArgs`로 되돌아간다.
- 이 `commands`는 **엑셀 추가기능 선언의 `commands`(docs/06, 단축키 명령 배열)와 다르다.** 추가기능 선언은 배열, 하위 프로그램 선언은 객체이고, 서로 읽지 않는다.
- `actions`(계약 2): `{ "clean": ["--clean"], "open-main": ["--main"] }` — 명령 id → 실행 인자(위 `settings`와 같은 규칙: 문자열 1~16개, 200자 이하, 최대 64개). 명령 id는 영문·숫자·`.`·`-`만(밑줄·공백 불가). 칼퇴는 같은 exe에 그 인자를 붙여 켜고, 설정창처럼 "켜짐"으로 세지 않는다. 틀린 항목은 그 항목만 무시.
- 새 명령은 이 객체에 키를 더하는 식으로만 늘린다(예: 나중에 `restart`). 칼퇴가 모르는 키는 무시되므로 먼저 낸 선언이 옛 칼퇴에서도 안전하다.

## 4. layout — 탭·버튼 힌트 (계약 2)
```json
"contract_version": 2,
"commands": { "actions": { "clean": ["--clean"], "open-main": ["--main"] } },
"layout": { "tabs": [
  { "id": "tools", "title": "도구", "kind": "buttons", "summary": "자주 쓰는 도면 정리",
    "groups": [ { "title": "도면", "buttons": [
      { "id": "a", "label": "선 정리", "tip": "겹친 선을 합쳐요", "icon": "icons/clean.png", "emphasis": "primary", "action": "clean" } ] } ] },
  { "id": "main", "title": "본 창", "kind": "window", "label": "열기", "action": "open-main" },
  { "id": "guide", "title": "도움말" },
  { "id": "packs", "hidden": true }
] }
```
- **아무것도 안 적으면**(layout 없음·틀림) 기본 탭 `[하위 모듈][찾기][팩][라이브러리][사용법][패치노트]`.
- **기본 탭**(id `modules`·`find`·`packs`·`libraries`·`guide`·`changes`): `title`로 이름, `hidden: true`로 숨김, 배열 순서로 순서를 바꾼다. **적지 않은 기본 탭은 기본 순서로 뒤에 붙는다**(칼퇴가 나중에 기본 탭을 더해도 숨겨지지 않는다). 모두 숨기면 기본으로 되돌린다.
- **사용자 정의 탭**: 기본 id가 아닌 id(영문·숫자·`.`·`-`).
  - `kind: "buttons"`: `groups[].buttons[]`. 버튼 = `id`(필수, 묶음 안 유일)·`label`(필수)·`action`(필수, `commands.actions`에 있어야 함)·`tip`·`icon`·`emphasis`("primary"면 강조색). `icon`은 묶음 안 안전한 상대 경로의 `.png`(그림은 파일로 둔다; 틀리면 아이콘 없이). `summary`는 탭 위 한 줄 안내.
  - `kind: "window"`(**탈출구**): 자기 창을 여는 탭. [열기] 단추 하나(`label`로 이름), `action` 필수. 복잡한 화면은 자기 창으로 열되, 창을 그릴 때는 **칼퇴와 같은 공통 컨트롤·디자인 토큰을 쓰는 것이 의무**다(§5).
  - 모르는 `kind`(앞으로 생길 `embed` 등)는 그 탭만 무시한다.
- **한도**: 탭 12개, 묶음 8개/탭, 버튼 24개/묶음, 탭 이름 24자, 버튼 이름 40자, 설명 120자. 넘는 것은 잘린다/무시된다.
- **순서와 보임**: 배열 순서 그대로. 같은 id는 첫 번째만.
- 칼퇴가 그림: 탭 줄(`TabStrip`), 구역 머리(`CatalogSectionHeader`), 공통 단추(`FlatButton`). 하위 프로그램은 색·글꼴·크기를 정할 수 없다.
- **module.json의 layout(예약)**: 모듈이 자기 버튼을 앱의 탭에 얹는 것(같은 id 탭에 묶음 추가)은 같은 모양으로 계약 2에 예약돼 있고 **아직 읽지 않는다**(모르는 칸이라 무시). 앱 선언(`<id>.app.json`)의 layout만 읽는다.

## 5. 탈출구와 디자인 토큰
- 칼퇴가 하위 프로그램을 켤 때(설정창·명령 포함) 환경변수 **`CLOCKOUT_THEME_FILE`**로 `<데이터 루트>\apps\<id>\theme.tokens.json`을 준다: `{"spec":0,"colors":{"window":"#F6F7F9","listBack","border","rowLine","groupBack","text","muted","accent","accentText","danger","warning"},"font":{"body","title","bodyPt","smallPt","titlePt"},"metrics":{"radius","controlHeight","scale"}}`. 키는 늘어도 지우거나 뜻을 바꾸지 않는다(읽는 쪽은 모르는 키 무시).
- 자기 창·패널을 그리는 하위 프로그램은 이 토큰으로 색·글꼴을 맞추고, 단추·목록 같은 컨트롤은 칼퇴 모양(평평한 둥근 사각형, 파랑 강조 하나)을 따라야 한다.

## 6. 칼퇴가 읽는 방법 (구현: `ClockOut.Core.Apps.ModuleContract`, `AppCommands`, `AppLayout`)
- `AppBundleManifest.ContractVersion`·`Commands`·`SettingsCommand`(우선순위 적용)·`Actions`·`Layout`·`LayoutTabs`(그릴 탭)·`Notes`(알림글).
- 화면: `AppDetailTabs`가 `LayoutTabs`대로 탭 줄을 짜고, 사용자 정의 탭은 `LayoutButtonsPage`, 버튼은 `AppProcesses.StartAction`.
- `ModuleEntry.ContractVersion`(module.json). module.json의 `spec != 0`만 거부한다.
- 알림글: 칼퇴보다 높은 계약 / 색·글꼴 같은 모양 칸(`theme`·`color(s)`·`style(s)`·`font(s)`·`ui`·`view`·`columns`) / 그 밖의 모르는 칸. 개발자 설치 때 상태줄에 보인다. "아는 칸"은 선언 클래스의 `JsonPropertyName`에서 자동으로 뽑는다.

## 7. 하위 프로그램 쪽에서 지킬 것
- 선언에는 위 표의 칸만 쓴다. 탭·버튼 구성은 layout으로 직접 정하고, 칼퇴 화면 모양이 바뀌어도 선언을 고치지 않는다.
- 칼퇴에 알리고 싶은 새 기능은 선언 칸이 아니라 **칼퇴 쪽 개편 요청**(bridge)으로 한다. 칼퇴가 칸을 추가하면 선택 항목이라 이미 낸 판은 그대로 동작한다.
- 환경변수(`CLOCKOUT_MODULES_FILE`·`CLOCKOUT_SETTINGS_FILE`·`CLOCKOUT_MODULES_DIR`)와 파일 형식(`spec` 0)도 같은 원칙: 필드를 더할 수는 있어도 지우거나 뜻을 바꾸지 않는다. **읽는 쪽도 모르는 필드를 무시**해야 한다.
