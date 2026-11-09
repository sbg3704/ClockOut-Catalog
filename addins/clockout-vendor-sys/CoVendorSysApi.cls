VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoVendorSysApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==============================================================
' 기능: clockout-vendor-sys의 공개 진입점 개체. 외부 클래스를 만들어(또는 팩토리 기본 개체를) As Object로 돌려준다.
'       stdVBA 클래스는 기본 개체가 팩토리라 그 개체를 그대로 돌려주고, 쓰는 쪽이 .Create...로 새 개체를 만든다.
' 사용법: Set v = CoVendorClient_Api("sys", "0.1.000", "내 추가기능") 다음 v.CoVendor_...()
' 의존: vendor.txt의 외부 원본
'==============================================================
Option Explicit

' stdClipboard: .Text(읽기·쓰기) .Clear 등
Public Function CoVendor_Clipboard() As Object
    Set CoVendor_Clipboard = stdClipboard
End Function

' stdPerformance: .CreateMeasure(이름) .CreateOptimiser()
Public Function CoVendor_Perf() As Object
    Set CoVendor_Perf = stdPerformance
End Function

' stdReg: .CreateFromKey("HKEY_CURRENT_USER\Software")
Public Function CoVendor_Reg() As Object
    Set CoVendor_Reg = stdReg
End Function

' stdProcess: .Create(명령) .CreateFromProcessId(번호)
Public Function CoVendor_Process() As Object
    Set CoVendor_Process = stdProcess
End Function

' stdWindow: .CreateFromApplication() .CreateFromHwnd(창)
Public Function CoVendor_Window() As Object
    Set CoVendor_Window = stdWindow
End Function

' stdAcc: .CreateFromApplication() .CreateFromPath(경로)
Public Function CoVendor_Acc() As Object
    Set CoVendor_Acc = stdAcc
End Function

' stdImage: 이미지 만들기·그리기·파일로 내보내기
Public Function CoVendor_Image() As Object
    Set CoVendor_Image = stdImage
End Function

' stdCOM: .Create(개체) 저수준 COM 호출
Public Function CoVendor_Com() As Object
    Set CoVendor_Com = stdCOM
End Function
