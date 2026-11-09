Attribute VB_Name = "CoVendorSys"
'==============================================================
' 기능: 칼퇴 외부 라이브러리 묶음 clockout-vendor-sys.xlam(시스템)의 바깥 입구. 이 모듈만 Option Private Module이 없어
'       다른 통합 문서(추가기능)가 Application.Run으로 부를 수 있다. 외부 원본은 vendor.txt에 적은 파일 그대로다.
' 사용법(추가기능 쪽, 연결부 addins/_client/CoVendorClient.bas):
'   Dim v As Object
'   Set v = CoVendorClient_Api("sys", "0.1.000", "내 추가기능")      ' 없거나 판이 낮으면 안내 후 Nothing
'   If v Is Nothing Then Exit Sub
'   Set w = v.CoVendor_Window().CreateFromApplication()
'   s = v.CoVendor_Clipboard().Text
'   직접 부를 때: Application.Run("'clockout-vendor-sys.xlam'!CoVendor_Version"), Application.Run("'clockout-vendor-sys.xlam'!CoVendor_Api")
' 팩토리(개체 CoVendorSysApi의 메서드, 모두 As Object로 늦은 바인딩):
'   CoVendor_Clipboard() : stdClipboard: .Text(읽기·쓰기) .Clear 등
'   CoVendor_Perf() : stdPerformance: .CreateMeasure(이름) .CreateOptimiser()
'   CoVendor_Reg() : stdReg: .CreateFromKey("HKEY_CURRENT_USER\Software")
'   CoVendor_Process() : stdProcess: .Create(명령) .CreateFromProcessId(번호)
'   CoVendor_Window() : stdWindow: .CreateFromApplication() .CreateFromHwnd(창)
'   CoVendor_Acc() : stdAcc: .CreateFromApplication() .CreateFromPath(경로)
'   CoVendor_Image() : stdImage: 이미지 만들기·그리기·파일로 내보내기
'   CoVendor_Com() : stdCOM: .Create(개체) 저수준 COM 호출
' 주의: 판을 올릴 때 PACK_VERSION과 clockout-vendor-sys.clockout.json의 version을 함께 고친다.
' 시험: Test_CoVendorSys
' 의존: CoVendorSysApi.cls와 vendor.txt의 외부 원본. clockout-lib와 추가기능은 참조하지 않는다.
'==============================================================
Option Explicit

Private Const PACK_VERSION As String = "0.1.000"
Private mApi As CoVendorSysApi

Public Function CoVendor_Version() As String
    CoVendor_Version = PACK_VERSION
End Function

Public Function CoVendor_Api() As Object
    If mApi Is Nothing Then Set mApi = New CoVendorSysApi
    Set CoVendor_Api = mApi
End Function

'--- 시험 (Excel 안에서 직접 실행: 보기 > 직접 실행 창) --------------------
' 다른 통합 문서처럼 Application.Run으로 개체를 얻어 늦은 바인딩으로 부른다.
Public Sub Test_CoVendorSys()
    Dim f As Long, api As Object
    Set api = Application.Run("'" & ThisWorkbook.Name & "'!CoVendor_Api")
    f = f + Chk("같은 개체", api Is CoVendor_Api())
    f = f + Chk("판", Application.Run("'" & ThisWorkbook.Name & "'!CoVendor_Version") = PACK_VERSION)
    f = f + Chk("stdWindow", Not api.CoVendor_Window().CreateFromApplication() Is Nothing)
    f = f + Chk("stdReg 읽기", Not api.CoVendor_Reg().CreateFromKey("HKEY_CURRENT_USER\Software") Is Nothing)
    f = f + Chk("stdPerformance", Not api.CoVendor_Perf() Is Nothing)
    f = f + Chk("stdClipboard", Not api.CoVendor_Clipboard() Is Nothing)
    f = f + Chk("stdProcess", Not api.CoVendor_Process() Is Nothing)
    f = f + Chk("stdAcc", Not api.CoVendor_Acc() Is Nothing)
    f = f + Chk("stdImage", Not api.CoVendor_Image() Is Nothing)
    f = f + Chk("stdCOM", Not api.CoVendor_Com() Is Nothing)
    If f = 0 Then
        Debug.Print "CoVendorSys 시험: 모두 통과"
        MsgBox "CoVendorSys 시험: 모두 통과", vbInformation, "칼퇴"
    Else
        MsgBox "CoVendorSys 시험: " & f & "개 실패 (직접 실행 창 확인)", vbExclamation, "칼퇴"
    End If
End Sub

Private Function Chk(ByVal testName As String, ByVal ok As Boolean) As Long
    If Not ok Then Debug.Print "[실패] " & testName: Chk = 1
End Function
