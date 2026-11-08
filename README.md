# ClockOut Catalog

칼퇴(ClockOut)가 검색·다운로드에 쓰는 추가기능 목록 저장소입니다. 운영자가 검수한 것만 올립니다.

- `catalog.json`: 추가기능과 팩 목록(id, 이름, 판, 설명, 다운로드 주소, SHA-256). 칼퇴 앱은 이 파일 하나를 받아 읽습니다.
- `addins/<id>/`: 선언 JSON(`<id>.clockout.json`), README, `*.bas` 소스, `ThisWorkbook_코드.txt`처럼 사람이 읽는 자료. 소스는 빌드 스크립트가 읽도록 폴더 바로 아래에 둔다(패키지 zip 안에서만 `src/`로 들어간다).
- `addins/_common/`: 모든 추가기능이 쓰는 공용 VBA 라이브러리(Co* 모듈). 사용법과 이름 규칙은 그 폴더의 README.
- `tools/build_addin.ps1`: Windows Excel에서 컴파일하고 `.xlam`과 설치용 zip을 만든다(`tools/build_addin.md`). `tools/vba_encode.py`: UTF-8 소스를 VBA 편집기용 CP949로 바꾼다. `tools/gen_colibapi.py`: 공용 lib의 공개 진입점 `CoLibApi.cls`를 만든다(lib 함수를 고치면 다시 돌린다).
- `vendor/vba/`: 외부 공개 VBA 라이브러리 원본(MIT만, 고치지 않음). 출처·커밋·라이선스는 `vendor/vba/SOURCES.md`. 전체를 `vba-vendor.zip` 한 묶음으로 Release에 올리고 `catalog.json`의 `libraries`에 등록한다(공용 lib와 별개).
- 바이너리(.xlam, 패키지·팩 zip)는 저장소에 커밋하지 않고 GitHub Release 파일로 올립니다.

목록 규격은 칼퇴 저장소의 `docs/03`에 있습니다. 규격이 바뀌면 `spec` 값이 올라갑니다.
