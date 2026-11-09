VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoVendorTextApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==============================================================
' 기능: clockout-vendor-text의 공개 진입점 개체. 외부 클래스를 만들어(또는 팩토리 기본 개체를) As Object로 돌려준다.
'       stdVBA 클래스는 기본 개체가 팩토리라 그 개체를 그대로 돌려주고, 쓰는 쪽이 .Create...로 새 개체를 만든다.
' 사용법: Set v = CoVendorClient_Api("text", "0.1.000", "내 추가기능") 다음 v.CoVendor_...()
' 의존: vendor.txt의 외부 원본, CoVendorRegex.cls
'==============================================================
Option Explicit

' 정규식 팩토리(CoVendorRegex, 순수 VBA 모드 고정): .Create(패턴, [플래그 "i" "m" "s" "g"]) → .Test .Match .MatchAll .Replace
Public Function CoVendor_Regex() As Object
    Set CoVendor_Regex = New CoVendorRegex
End Function

' HTML 팩토리(stdHTML): .Create(태그) → .CreateChild .CreateLiteral .ToString
Public Function CoVendor_Html() As Object
    Set CoVendor_Html = stdHTML
End Function
