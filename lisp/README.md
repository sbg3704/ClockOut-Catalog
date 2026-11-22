# 리습(AutoCAD) 기준 모음

캐드선수(`seonsu`) 하위 모듈로 넣는 AutoCAD 리습의 기준만 둔다. **리습 소스는 여기 두지 않는다**(원본 자리: `sbg3704/ClockOut-Lisp`의 `modules/lisp/<모듈id>/`). 엑셀 추가기능의 `addins/`와 같은 역할이지만 소스가 다른 저장소에 있다.

| 문서 | 내용 |
|---|---|
| [리습_작성_기준.md](리습_작성_기준.md) | 만들기 전 지침, 폴더·선언(`module.json`), 코드 규칙, 공용 라이브러리와 **라이브러리 안내 필수 항목**, 사용자 안내, 점검표 |

## 공용 라이브러리
| id | 판 | 안내(함수·명령·활용법) |
|---|---|---|
| (아직 없음) | | 만들면 `ClockOut-Lisp/modules/lisp/<id>/README.md`를 여기서 가리킨다 |

## 카탈로그에 올릴 때
- 모듈은 캐드선수 빌드가 만든 `out/modules/catalog-modules.json`을 `catalog.json`의 `modules` 칸에 붙인다(`app: "seonsu"`). 순서는 ClockOut-Lisp `docs/03_릴리스_가이드.md`.
