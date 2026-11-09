VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoVendorDataApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==============================================================
' 기능: clockout-vendor-data의 공개 진입점 개체. 외부 클래스를 만들어(또는 팩토리 기본 개체를) As Object로 돌려준다.
'       stdVBA 클래스는 기본 개체가 팩토리라 그 개체를 그대로 돌려주고, 쓰는 쪽이 .Create...로 새 개체를 만든다.
' 사용법: Set v = CoVendorClient_Api("data", "0.1.000", "내 추가기능") 다음 v.CoVendor_...()
' 의존: vendor.txt의 외부 원본, CoVendorJson.cls
'==============================================================
Option Explicit

' JSON·날짜 도구(CoVendorJson): .Parse(글) .Stringify(값, [들여쓰기]) .ParseIso .ToIso .ToUtc .FromUtc
Public Function CoVendor_Json() As Object
    Set CoVendor_Json = New CoVendorJson
End Function

' 빈 사전(VBA-Dictionary): .Add 키, 값 / .Exists(키) / d(키)
Public Function CoVendor_Dict() As Object
    Set CoVendor_Dict = New Dictionary
End Function
