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

## 외부 라이브러리 묶음 연결부 `CoVendorClient.bas`
외부 공개 VBA(VBA-JSON, stdVBA 등)는 층·형식별 라이브러리 .xlam 4개로 따로 설치한다. 추가기능은 소스를 복사하지 않고 필요한 묶음만 부른다. 선언 `dependencies`에 `clockout-vendor-*`가 있으면 `build_addin.ps1`이 이 파일을 자동으로 넣는다. CoLibClient와 함께 넣어도 이름이 겹치지 않는다.

```vb
Dim v As Object, j As Object
Set v = CoVendorClient_Api("data", "0.1.000", "JSON 예제")   ' 묶음: data·text·array·sys (또는 전체 id)
If v Is Nothing Then Exit Sub                                 ' 없거나 판이 낮으면 안내 후 Nothing (네 번째 인자 quiet:=True면 조용히)
Set j = v.CoVendor_Json()
Set o = j.Parse(cellText): s = j.Stringify(o, 2)
```

| 묶음 id | 층 | 팩토리(진입점 개체의 메서드) | 원본 |
|---|---|---|---|
| `clockout-vendor-data` | 형식 | `CoVendor_Json`(Parse·Stringify·ISO/UTC 날짜), `CoVendor_Dict` | VBA-JSON, VBA-Dictionary |
| `clockout-vendor-text` | 글자 | `CoVendor_Regex`(순수 VBA 정규식), `CoVendor_Html` | stdRegex, stdHTML |
| `clockout-vendor-array` | 배열 | `CoVendor_Arr`(BetterArray), `CoVendor_StdArray`, `CoVendor_Enum`, `CoVendor_Lambda`, `CoVendor_Callback` | VBA-Better-Array, stdVBA |
| `clockout-vendor-sys` | 시스템 | `CoVendor_Clipboard`, `CoVendor_Perf`, `CoVendor_Reg`, `CoVendor_Process`, `CoVendor_Window`, `CoVendor_Acc`, `CoVendor_Image`, `CoVendor_Com` | stdVBA |

- 선언에는 쓰는 묶음만 `{"id": "clockout-vendor-data", "min_version": "0.1.000", "reason": "JSON 읽기·쓰기"}`처럼 적는다. 앱이 그 묶음만 받아 켠다(안 쓰는 묶음은 설치되지 않음 = 켜고 끄기).
- stdVBA 개체(정규식, 람다 등)는 팩토리 기본 개체를 돌려주므로 `.Create(...)`로 새 개체를 만든다. 람다처럼 다른 개체에 넘기는 것은 **같은 묶음에서 만든 것**을 넘긴다.
- 묶음끼리는 서로 부르지 않는다. 바닥 파일(stdICallable 등)은 필요한 묶음마다 원본 그대로 들어 있다.
- 예제: `addins/vendor-sample`(JSON 보기 좋게).
