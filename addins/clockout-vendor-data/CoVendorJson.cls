VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoVendorJson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================
' 기능: VBA-JSON(JsonConverter.bas)의 공개 함수를 개체 메서드로 감싼다. 표준 모듈 함수는 다른 통합 문서에서
'       늦은 바인딩으로 부를 수 없어서 이 개체를 거친다. 동작은 원본과 같다(오류도 원본처럼 던진다).
' 사용법:
'   Set j = v.CoVendor_Json()
'   Set o = j.Parse(jsonText)            ' {} -> Dictionary, [] -> Collection(1부터), null -> Null
'   s = j.Stringify(o)                   ' 한 줄. j.Stringify(o, 2)는 2칸 들여쓰기, j.Stringify(o, vbTab)도 된다
'   s = j.ToIso(Now) / dt = j.ParseIso("2026-11-09T06:06:00.000Z")   ' 지역 시간 <-> ISO 8601(UTC)
'   dt = j.ToUtc(Now) / dt = j.FromUtc(utc)
'   j.UseDoubleForLargeNumbers = True   ' 15자리 넘는 수를 글 대신 Double로(원본 JsonOptions)
' 의존: JsonConverter.bas, Dictionary.cls
'==============================================================
Option Explicit

' JSON 글을 개체(Dictionary 또는 Collection)로. 형식이 틀리면 원본처럼 오류(10001)를 던진다.
Public Function Parse(ByVal jsonText As String) As Object
    Set Parse = JsonConverter.ParseJson(jsonText)
End Function

' 값(Dictionary·Collection·배열·글·수·참거짓·Null·날짜)을 JSON 글로. whitespace는 들여쓰기 칸 수나 글.
Public Function Stringify(ByVal jsonValue As Variant, Optional ByVal whitespace As Variant) As String
    Stringify = JsonConverter.ConvertToJson(jsonValue, whitespace)
End Function

' ISO 8601 글을 지역 시간으로.
Public Function ParseIso(ByVal isoText As String) As Date
    ParseIso = JsonConverter.ParseIso(isoText)
End Function

' 지역 시간을 ISO 8601(UTC, 끝이 Z) 글로.
Public Function ToIso(ByVal localDate As Date) As String
    ToIso = JsonConverter.ConvertToIso(localDate)
End Function

' 지역 시간 -> UTC.
Public Function ToUtc(ByVal localDate As Date) As Date
    ToUtc = JsonConverter.ConvertToUtc(localDate)
End Function

' UTC -> 지역 시간.
Public Function FromUtc(ByVal utcDate As Date) As Date
    FromUtc = JsonConverter.ParseUtc(utcDate)
End Function

Public Property Get UseDoubleForLargeNumbers() As Boolean
    UseDoubleForLargeNumbers = JsonConverter.JsonOptions.UseDoubleForLargeNumbers
End Property
Public Property Let UseDoubleForLargeNumbers(ByVal flag As Boolean)
    JsonConverter.JsonOptions.UseDoubleForLargeNumbers = flag
End Property

Public Property Get AllowUnquotedKeys() As Boolean
    AllowUnquotedKeys = JsonConverter.JsonOptions.AllowUnquotedKeys
End Property
Public Property Let AllowUnquotedKeys(ByVal flag As Boolean)
    JsonConverter.JsonOptions.AllowUnquotedKeys = flag
End Property

Public Property Get EscapeSolidus() As Boolean
    EscapeSolidus = JsonConverter.JsonOptions.EscapeSolidus
End Property
Public Property Let EscapeSolidus(ByVal flag As Boolean)
    JsonConverter.JsonOptions.EscapeSolidus = flag
End Property
