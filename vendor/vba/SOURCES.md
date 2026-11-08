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

## 묶음
이 폴더 전체를 `vba-vendor.zip` 하나로 묶어 Release에 올리고 `catalog.json`의 `libraries`에 SHA-256과 함께 등록한다. 공용 lib(`addins/_common`, clockout-lib)와는 별개 묶음이다. 빌드 스크립트(`tools/build_addin.ps1`)와 `tools/vba_encode.py`는 `addins/`만 보므로 이 폴더를 자동으로 넣지 않는다. 쓸 때는 필요한 파일만 추가기능 폴더로 복사한다.
