VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoVendorArrayApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==============================================================
' 기능: clockout-vendor-array의 공개 진입점 개체. 외부 클래스를 만들어(또는 팩토리 기본 개체를) As Object로 돌려준다.
'       stdVBA 클래스는 기본 개체가 팩토리라 그 개체를 그대로 돌려주고, 쓰는 쪽이 .Create...로 새 개체를 만든다.
' 사용법: Set v = CoVendorClient_Api("array", "0.1.000", "내 추가기능") 다음 v.CoVendor_...()
' 의존: vendor.txt의 외부 원본
'==============================================================
Option Explicit

' 새 BetterArray: .Push .Sort .Item(i) .Length .ToString
Public Function CoVendor_Arr() As Object
    Set CoVendor_Arr = New BetterArray
End Function

' stdArray 팩토리: .Create(값들) / .CreateFromArray(배열) → .Map .Filter .Join
Public Function CoVendor_StdArray() As Object
    Set CoVendor_StdArray = stdArray
End Function

' stdEnumerator 팩토리: .CreateFromArray(배열)
Public Function CoVendor_Enum() As Object
    Set CoVendor_Enum = stdEnumerator
End Function

' stdLambda 팩토리: .Create("$1*2") → .Run(값)
Public Function CoVendor_Lambda() As Object
    Set CoVendor_Lambda = stdLambda
End Function

' stdCallback 팩토리: .CreateFromModule(모듈, 함수)
Public Function CoVendor_Callback() As Object
    Set CoVendor_Callback = stdCallback
End Function
