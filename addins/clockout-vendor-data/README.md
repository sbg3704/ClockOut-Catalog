# 칼퇴 외부 라이브러리: 형식(JSON·사전·날짜) (`clockout-vendor-data`)

외부 공개 VBA 라이브러리를 원본 그대로 컴파일한 라이브러리 .xlam입니다. 추가기능이 소스를 복사해 넣지 않고 설치된 이 파일을 불러 씁니다.

- 판: 0.1.000 · 종류: 라이브러리(`type: library`) · 라이선스: MIT (VBA-Dictionary, VBA-JSON의 LICENSE 전문을 zip의 `licenses/`에 함께 넣음)
- 들어 있는 원본: `Dictionary.cls`, `JsonConverter.bas` (출처·커밋: `vendor/vba/SOURCES.md`)

## 쓰는 법
```vb
Dim v As Object
Set v = CoVendorClient_Api("data", "0.1.000", "내 추가기능")   ' 없거나 판이 낮으면 안내 후 Nothing
If v Is Nothing Then Exit Sub
Set j = v.CoVendor_Json()
Set o = j.Parse("{""a"":[1,2]}")      ' 개체는 Dictionary, 배열은 Collection
s = j.Stringify(o, 2)                    ' 들여쓰기 2칸
```
추가기능 선언에는 `"dependencies": [{"id": "clockout-vendor-data", "min_version": "0.1.000", "reason": "..."}]`를 적습니다.

| 메서드 | 돌려주는 것 |
|---|---|
| `CoVendor_Json()` | JSON·날짜 도구(CoVendorJson): .Parse(글) .Stringify(값, [들여쓰기]) .ParseIso .ToIso .ToUtc .FromUtc |
| `CoVendor_Dict()` | 빈 사전(VBA-Dictionary): .Add 키, 값 / .Exists(키) / d(키) |
