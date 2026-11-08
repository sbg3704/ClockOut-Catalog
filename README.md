# ClockOut Catalog

칼퇴(ClockOut)가 검색·다운로드에 쓰는 추가기능 목록 저장소입니다. 운영자가 검수한 것만 올립니다.

- `catalog.json`: 추가기능과 팩 목록(id, 이름, 판, 설명, 다운로드 주소, SHA-256). 칼퇴 앱은 이 파일 하나를 받아 읽습니다.
- `addins/<id>/`: 선언 JSON, README, `src/*.bas`처럼 사람이 읽는 자료.
- 바이너리(.xlam, 패키지·팩 zip)는 저장소에 커밋하지 않고 GitHub Release 파일로 올립니다.

목록 규격은 칼퇴 저장소의 `docs/03`에 있습니다. 규격이 바뀌면 `spec` 값이 올라갑니다.
