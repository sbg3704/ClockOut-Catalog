# 외부 공개 VBA 라이브러리 출처 (vendor/vba)

칼퇴 매크로를 만들 때 코드를 새로 짜지 않도록 받아 둔 공개 라이브러리 원본이다. 받은 날: 2026-11-08(계정2). 원본에서 **한 글자도 고치지 않았다**(바이트 단위로 원본 커밋과 같은지 비교함). 고쳐 써야 하면 원본은 두고 `addins/_common/`에 감싸는 모듈을 만들거나, 바꾼 사본을 따로 두고 이 표에 적는다.

## 받은 것
| 폴더 | 원본 주소 | 받은 커밋 | 커밋 날짜 | 받은 날짜 | 라이선스 | 작성자 | 수정 여부 | 받은 파일 |
|---|---|---|---|---|---|---|---|---|
| `VBA-JSON` | https://github.com/VBA-tools/VBA-JSON | `06b6c45a2e8086a7a761f00bead7460852d6b602` | 2019-01-28 | 2026-11-08 | MIT | Tim Hall | 없음 | `JsonConverter.bas`, `LICENSE`, `README.md` |
| `VBA-Dictionary` | https://github.com/VBA-tools/VBA-Dictionary | `757aea9207648821dfc0c382c4a5431836749c3e` | 2020-01-22 | 2026-11-08 | MIT | Tim Hall | 없음 | `Dictionary.cls`, `LICENSE`, `README.md` |
| `VBA-UTC` | https://github.com/VBA-tools/VBA-UTC | `122928ec2062e13a73333d3d00f3a8cc73f04304` | 2019-01-28 | 2026-11-08 | MIT | Tim Hall | 없음 | `UtcConverter.bas`, `LICENSE`, `README.md` |
| `VBA-Web` | https://github.com/VBA-tools/VBA-Web | `9dbcc751d177099f20c96c5ee332ec10ef47423c` | 2019-10-04 | 2026-11-08 | MIT | Tim Hall | 없음 | `src/`, `authenticators/`, `LICENSE`, `README.md` |
| `VBA-Better-Array` | https://github.com/Senipah/VBA-Better-Array | `b69de38867d16415cd07cec68ca97658bf6f6e3c` | 2026-04-10 | 2026-11-08 | MIT | Nick Morgan (Senipah) | 없음 | `src/`, `LICENSE`, `README.md` |
| `stdVBA` | https://github.com/sancarn/stdVBA | `230aea3aef43cd4bb78a728e7a18f70e11b9a77b` | 2026-11-01 | 2026-11-08 | MIT | James Warren (sancarn) | 없음 | `src/` 맨 위 파일(`src/WIP/` 제외), `LICENSE`, `README.md` |

- 각 폴더에 원본 `LICENSE`가 그대로 있다. **이 코드를 넣은 추가기능을 배포할 때는 그 라이브러리의 `LICENSE` 전문(저작권 줄 포함)을 zip이나 README에 함께 넣는다**(MIT 조건).
- 원본 저장소의 `.xlsm` 예제·설치 파일, `docs/`, `specs/`·`tests/`는 받지 않았다(바이너리이거나 원본 소스가 아님).
- 쓸 곳: VBA-JSON(중첩 JSON), VBA-Dictionary(Mac 겸용 Dictionary), VBA-UTC(UTC·ISO 8601), VBA-Web(HTTP·인증), VBA-Better-Array(배열 정렬·검색), stdVBA(정규식·클립보드·HTTP·람다 등).

## 받지 않은 것
| 라이브러리 | 원본 주소 | 확인한 커밋 | 라이선스 | 작성자 | 받지 않은 이유 |
|---|---|---|---|---|---|
| VBA-CSV-interface | https://github.com/ws-garcia/VBA-CSV-interface | `1e4ec36a15d7e9efe3edb93c36fecff86c8fc825` | GPL-3.0 | W. García (ws-garcia) | GPL은 이 코드를 넣은 추가기능 전체를 같은 GPL로 공개해야 하는 조건이 있어 칼퇴의 공개 배포·검수 방식과 맞지 않음 |
| VBA-Expressions | https://github.com/ws-garcia/VBA-Expressions | `3ffee6f2e173364a27a48f2dc657c23ee6fb33a4` | GPL-3.0 | W. García (ws-garcia) | 위와 같음(필요하면 `CoNum_EvalSafe`를 lib에서 직접 만든다) |

그 밖에 라이선스가 분명하지 않은 웹 코드 모음(Chip Pearson 등)과 VBA 편집기 도구인 Rubberduck(GPL, 라이브러리 아님)도 받지 않았다.

## 묶음 (2026-11-09부터: 라이브러리 .xlam 4개)
이 폴더의 원본을 **고치지 않고** 층·형식별 라이브러리 .xlam으로 컴파일해 따로 배포한다. 각 묶음 폴더(`addins/clockout-vendor-*/`)의 `vendor.txt`가 넣을 원본 파일을 정하고, `tools/build_addin.ps1`이 그 파일을 함께 컴파일하며 zip의 `licenses/`에 해당 `LICENSE` 전문과 이 표를 넣는다.

| 묶음 | 넣은 원본 | 뺀 것과 이유 |
|---|---|---|
| `clockout-vendor-data` | VBA-Dictionary `Dictionary.cls`, VBA-JSON `JsonConverter.bas` | VBA-UTC `UtcConverter.bas`: JsonConverter가 같은 이름의 공개 함수(ParseUtc·ConvertToUtc·ParseIso·ConvertToIso)를 이미 가지고 있어 중복 선언이 됨. 날짜 변환은 JsonConverter 쪽을 씀 |
| `clockout-vendor-text` | stdVBA `stdRegex.cls`, `stdHTML.cls` | |
| `clockout-vendor-array` | VBA-Better-Array `BetterArray.cls`, stdVBA `stdICallable`·`stdArray`·`stdEnumerator`·`stdLambda`·`stdCallback` | Better-Array의 시험·예제 파일(`TestModule_*`, `ValuesList_*`, `ArrayGenerator` 등): 라이브러리 본체가 아님 |
| `clockout-vendor-sys` | stdVBA `stdICallable`·`stdClipboard`·`stdPerformance`·`stdReg`·`stdProcess`·`stdAcc`·`stdWindow`·`stdImage`·`stdCOM` | `stdTimer`: 'VBA 프로젝트 개체 모델 접근 신뢰'와 숨김 통합 문서가 필요함. `stdUIElement`: MSForms 참조가 필요함(사용자 정의 폼 전용). `stdHTTP`·`stdWebSocket`·`stdWebView`·`stdFiber`·`stdSentry`·`stdRibbon`: 네트워크·실험 기능이라 다음 판에 따로 논의 |
| (다음 판) | VBA-Web | `WebHelpers.bas`가 JSON·UTC 공개 함수 6개를 같은 이름으로 또 선언함. 네트워크 기능은 오프라인 원칙과 함께 따로 정함 |

- 의존 관계(주석을 뺀 코드 기준): 바닥은 `stdICallable`과 `Dictionary`이고 나머지는 이 둘에만 기대거나 혼자 돈다. `stdWindow`는 `stdAcc`에 기댄다. 묶음끼리는 서로 부르지 않는다.
- CP949에 없는 글자(예: `stdImage.cls` 주석의 é)는 컴파일용 사본에서만 `?`로 바꾼다. 이 폴더의 원본은 그대로다.
- 예전 소스 모음 `vba-vendor.zip`(`catalog.json`의 `libraries`)은 묶음 4개를 올리면서 카탈로그에서 뺀다.
