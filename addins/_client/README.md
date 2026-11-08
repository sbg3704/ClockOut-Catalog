# 공용 라이브러리 클라이언트 (`addins/_client`)

`CoLibClient.bas` — 따로 설치하는 공용 라이브러리(clockout-lib.xlam)를 추가기능이 부르는 얇은 연결부. 선언 JSON의 `dependencies`에 `clockout-lib`이 있고 `bundled`가 아니면 `tools/build_addin.ps1`이 이 파일을 추가기능에 자동으로 넣는다(`bundled: true`이거나 선언이 없으면 예전처럼 `addins/_common` 소스를 함께 컴파일).

```vb
Dim coApi As Object
Set coApi = CoLibClient_Api("0.1.001", "십자 음영")   ' 없거나 판이 낮으면 안내 후 Nothing
If coApi Is Nothing Then Exit Sub
coApi.CoGuard_Notice "안내 글"                        ' 늦은 바인딩: 메서드 이름은 라이브러리 함수 이름 그대로
```
- 이벤트처럼 자주 불리는 곳은 세 번째 인자 `quiet:=True`로 안내 창을 끈다.
- 추가기능 선언의 `dependencies.min_version`과 코드의 최소 판을 같게 적는다.
- 라이브러리가 아직 안 켜졌을 수 있으니, 시작 때 라이브러리를 부르는 코드는 `Workbook_Open`에서 `Application.OnTime`으로 조금 뒤에 실행한다(십자 음영·선택 요약의 `ThisWorkbook_코드.txt` 참고).
