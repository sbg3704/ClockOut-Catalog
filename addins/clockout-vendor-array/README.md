# 칼퇴 외부 라이브러리: 배열(BetterArray·stdArray·람다) (`clockout-vendor-array`)

외부 공개 VBA 라이브러리를 원본 그대로 컴파일한 라이브러리 .xlam입니다. 추가기능이 소스를 복사해 넣지 않고 설치된 이 파일을 불러 씁니다.

- 판: 0.1.000 · 종류: 라이브러리(`type: library`) · 라이선스: MIT (VBA-Better-Array, stdVBA의 LICENSE 전문을 zip의 `licenses/`에 함께 넣음)
- 들어 있는 원본: `BetterArray.cls`, `stdICallable.cls`, `stdArray.cls`, `stdEnumerator.cls`, `stdLambda.cls`, `stdCallback.cls` (출처·커밋: `vendor/vba/SOURCES.md`)

## 쓰는 법
```vb
Dim v As Object
Set v = CoVendorClient_Api("array", "0.1.000", "내 추가기능")   ' 없거나 판이 낮으면 안내 후 Nothing
If v Is Nothing Then Exit Sub
Set a = v.CoVendor_Arr(): a.Push 3, 1, 2: a.Sort
Set dbl = v.CoVendor_Lambda().Create("$1*2")
s = v.CoVendor_StdArray().Create(1, 2, 3).Map(dbl).Join(",")   ' "2,4,6" (람다는 같은 묶음에서 만든 것을 넘긴다)
```
추가기능 선언에는 `"dependencies": [{"id": "clockout-vendor-array", "min_version": "0.1.000", "reason": "..."}]`를 적습니다.

| 메서드 | 돌려주는 것 |
|---|---|
| `CoVendor_Arr()` | 새 BetterArray: .Push .Sort .Item(i) .Length .ToString |
| `CoVendor_StdArray()` | stdArray 팩토리: .Create(값들) / .CreateFromArray(배열) → .Map .Filter .Join |
| `CoVendor_Enum()` | stdEnumerator 팩토리: .CreateFromArray(배열) |
| `CoVendor_Lambda()` | stdLambda 팩토리: .Create("$1*2") → .Run(값) |
| `CoVendor_Callback()` | stdCallback 팩토리: .CreateFromModule(모듈, 함수) |
