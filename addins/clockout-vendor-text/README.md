# 칼퇴 외부 라이브러리: 글자(정규식·HTML) (`clockout-vendor-text`)

외부 공개 VBA 라이브러리를 원본 그대로 컴파일한 라이브러리 .xlam입니다. 추가기능이 소스를 복사해 넣지 않고 설치된 이 파일을 불러 씁니다.

- 판: 0.1.000 · 종류: 라이브러리(`type: library`) · 라이선스: MIT (stdVBA의 LICENSE 전문을 zip의 `licenses/`에 함께 넣음)
- 들어 있는 원본: `stdRegex.cls`, `stdHTML.cls` (출처·커밋: `vendor/vba/SOURCES.md`)

## 쓰는 법
```vb
Dim v As Object
Set v = CoVendorClient_Api("text", "0.1.000", "내 추가기능")   ' 없거나 판이 낮으면 안내 후 Nothing
If v Is Nothing Then Exit Sub
Set rx = v.CoVendor_Regex().Create("(\d+)-(\d+)")
If rx.Test(s) Then s = rx.Replace(s, "$2-$1")
```
추가기능 선언에는 `"dependencies": [{"id": "clockout-vendor-text", "min_version": "0.1.000", "reason": "..."}]`를 적습니다.

| 메서드 | 돌려주는 것 |
|---|---|
| `CoVendor_Regex()` | 정규식 팩토리(CoVendorRegex, 순수 VBA 모드 고정): .Create(패턴, [플래그 "i" "m" "s" "g"]) → .Test .Match .MatchAll .Replace |
| `CoVendor_Html()` | HTML 팩토리(stdHTML): .Create(태그) → .CreateChild .CreateLiteral .ToString |

## 주의
- stdRegex 최신판(2026-11-01 커밋)은 기본으로 실험용 기계어 실행 모드(useNative=True)를 씁니다. Excel이 꺼진 일이 있어서(2026-11-09 시험) `CoVendor_Regex().Create`는 항상 순수 VBA 모드로 만듭니다. 원본 파일은 고치지 않았습니다.
