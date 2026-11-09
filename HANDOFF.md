# HANDOFF

- 마지막 갱신: 2026-11-09
- 환경: office-host (`D:\CO\ClockOut-Catalog`)
- 계정: 2 (sbg3704)
- 브랜치: claude/docs-workspace-onboarding (origin/main 664af4e 기준)

## 완료
- 추가기능 11개 Release 등록(공용 lib·image-fit 해시 반영), lib를 `libraries`로 등록(PR#18까지 main에 합쳐짐).
- claude-workspace 체계 편입(CLAUDE.md, HANDOFF.md).

## 진행 중
- 없음(이 저장소 자체의 미완료 작업은 확인된 것 없음).

## 다음 할 일
- 앱 저장소 `HANDOFF.md`의 배포 순서 참고: 앱 릴리스(`app-v0.1.000`)는 이 저장소 Release로 올리며 **사용자가 직접** 만든다.
- 새 매크로·lib 요청은 앱 저장소 `협업/요청/`으로 받는다.

## 결정사항과 이유
- 바이너리는 커밋하지 않고 Release 파일로(저장소 크기·공개 저장소 관리).
- 공용 lib 판은 공개 함수 추가·동작 변경으로 Release를 다시 올릴 때만 올린다.

## 주의점 / 하지 말 것
- 공개 저장소: 비밀값·개인정보 금지. 사용자 PC 키보드·화면 조작 금지.
- 이전 로컬 브랜치 `claude/catalog-publish-9`는 기존 작업용이라 재사용하지 않음.
