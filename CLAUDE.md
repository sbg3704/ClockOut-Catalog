# ClockOut-Catalog

## 개요
칼퇴(ClockOut, 앱 저장소 `sbg3704/ClockOut`)가 검색·다운로드에 쓰는 **공개** 저장소. 검수된 추가기능 소스(`addins/`), 공용 VBA 라이브러리(`addins/_common/`, id `clockout-lib`), 외부 MIT 라이브러리 원본(`vendor/vba/`), 카탈로그 목록(`catalog.json`), 빌드 도구(`tools/`)가 있다. 공개 저장소이므로 비밀값·개인정보는 절대 올리지 않는다.
앱 쪽 규칙·판 번호는 앱 저장소 `docs/02_규칙.md`, 목록 규격은 앱 저장소 `docs/03`이 기준.

## 빌드 / 테스트
Windows Excel 데스크톱판 필요(VBA 프로젝트 개체 모델 접근 신뢰를 사용자가 한 번 켜 둠).
```powershell
powershell -ExecutionPolicy Bypass -File tools\build_addin.ps1 -Addin <id> [-RunTests]
powershell -ExecutionPolicy Bypass -File tools\build_addin.ps1 -Addin _common   # 공용 lib만
python tools\gen_colibapi.py   # lib 공개 함수를 고치면 CoLibApi.cls 다시 생성
```
자세한 절차: `tools/build_addin.md`. 결과물은 `out\`(git 제외).

## 폴더 구조
- `addins/<id>/` 선언 JSON(`<id>.clockout.json`)·README·`*.bas` 소스, `addins/_common/` 공용 lib
- `vendor/vba/` 외부 라이브러리 원본(MIT만, 고치지 않음, 출처 `SOURCES.md`)
- `catalog.json` 앱이 읽는 목록, `tools/` 빌드·배포 스크립트

## 규칙
- 작업 시작 시 HANDOFF.md를 읽고, 끝낼 때 갱신 후 커밋·push.
- 빌드는 반드시 **컴파일 확인**을 거치고 소스는 **CP949**로 변환(`tools/vba_encode.py`, 빌드 스크립트가 자동 처리)해 `.xlam`·설치용 zip을 만든다.
- **모듈화**: 파일 하나는 한 가지 책임, 300줄 안팎(800줄 넘기지 않음). 모든 소스 맨 위에 기능·사용법·의존을 한국어 주석으로.
- **라이브러리 원칙**: 추가기능 → 공용 lib 한 방향 의존. 같은 일을 하는 함수를 매크로마다 새로 만들지 말고 lib에서 찾아 쓰고, 없으면 lib에 더한다(공개 함수 이름·인자는 바꾸지 않음). 외부 라이브러리는 MIT만.
- **의존성 선언 형식**: `dependencies: [{id, min_version, reason}]`(번들이면 `bundled: true`).
- 판 번호는 앱 저장소 코디네이터만 매긴다. 작업 중에는 무게(+1/+10/+100)만 적는다.
- 자기 `claude/…` 브랜치에만 push. 사용자와는 항상 한국어, 결론부터 짧게.

## 하지 말 것
- **사용자 PC에서 키보드 입력·화면 조작 금지.** Excel은 보이지 않게 띄우고, 스크립트가 띄운 Excel만 종료한다.
- 바이너리(.xlam, zip)를 커밋하지 않는다(Release 파일로). **공개 Release 생성은 사용자가 직접.**
- force push·이력 다시 쓰기 금지, 비밀값(토큰·키) 커밋 금지.
- 설계로(Web_PMIS_GIS) 서버 작업은 다른 프로젝트: 요청서 `.md`로 전달.
- 다른 계정·세션을 자동 호출하는 구조 만들기 금지.
