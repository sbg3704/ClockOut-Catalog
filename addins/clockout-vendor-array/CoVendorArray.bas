Attribute VB_Name = "CoVendorArray"
'==============================================================
' 기능: 칼퇴 외부 라이브러리 묶음 clockout-vendor-array.xlam(배열)의 바깥 입구. 이 모듈만 Option Private Module이 없어
'       다른 통합 문서(추가기능)가 Application.Run으로 부를 수 있다. 외부 원본은 vendor.txt에 적은 파일 그대로다.
' 사용법(추가기능 쪽, 연결부 addins/_client/CoVendorClient.bas):
'   Dim v As Object
'   Set v = CoVendorClient_Api("array", "0.1.000", "내 추가기능")      ' 없거나 판이 낮으면 안내 후 Nothing
'   If v Is Nothing Then Exit Sub
'   Set a = v.CoVendor_Arr(): a.Push 3, 1, 2: a.Sort
'   Set dbl = v.CoVendor_Lambda().Create("$1*2")
'   s = v.CoVendor_StdArray().Create(1, 2, 3).Map(dbl).Join(",")   ' "2,4,6" (람다는 같은 묶음에서 만든 것을 넘긴다)
'   직접 부를 때: Application.Run("'clockout-vendor-array.xlam'!CoVendor_Version"), Application.Run("'clockout-vendor-array.xlam'!CoVendor_Api")
' 팩토리(개체 CoVendorArrayApi의 메서드, 모두 As Object로 늦은 바인딩):
'   CoVendor_Arr() : 새 BetterArray: .Push .Sort .Item(i) .Length .ToString
'   CoVendor_StdArray() : stdArray 팩토리: .Create(값들) / .CreateFromArray(배열) → .Map .Filter .Join
'   CoVendor_Enum() : stdEnumerator 팩토리: .CreateFromArray(배열)
'   CoVendor_Lambda() : stdLambda 팩토리: .Create("$1*2") → .Run(값)
'   CoVendor_Callback() : stdCallback 팩토리: .CreateFromModule(모듈, 함수)
' 주의: 판을 올릴 때 PACK_VERSION과 clockout-vendor-array.clockout.json의 version을 함께 고친다.
' 시험: Test_CoVendorArray
' 의존: CoVendorArrayApi.cls와 vendor.txt의 외부 원본. clockout-lib와 추가기능은 참조하지 않는다.
'==============================================================
Option Explicit

Private Const PACK_VERSION As String = "0.1.000"
Private mApi As CoVendorArrayApi

Public Function CoVendor_Version() As String
    CoVendor_Version = PACK_VERSION
End Function

Public Function CoVendor_Api() As Object
    If mApi Is Nothing Then Set mApi = New CoVendorArrayApi
    Set CoVendor_Api = mApi
End Function

'--- 시험 (Excel 안에서 직접 실행: 보기 > 직접 실행 창) --------------------
' 다른 통합 문서처럼 Application.Run으로 개체를 얻어 늦은 바인딩으로 부른다.
Public Sub Test_CoVendorArray()
    Dim f As Long, api As Object
    Dim a As Object, lam As Object
    Set api = Application.Run("'" & ThisWorkbook.Name & "'!CoVendor_Api")
    f = f + Chk("같은 개체", api Is CoVendor_Api())
    f = f + Chk("판", Application.Run("'" & ThisWorkbook.Name & "'!CoVendor_Version") = PACK_VERSION)
    Set a = api.CoVendor_Arr()
    a.Push 3, 1, 2
    a.Sort
    f = f + Chk("BetterArray 정렬", a.Item(a.LowerBound) = 1 And a.Length = 3)
    Set lam = api.CoVendor_Lambda().Create("$1*2")
    f = f + Chk("람다", lam.Run(21) = 42)
    f = f + Chk("stdArray Map", api.CoVendor_StdArray().Create(1, 2, 3).Map(lam).Join(",") = "2,4,6")
    f = f + Chk("stdArray 길이", api.CoVendor_StdArray().CreateFromArray(Array(5, 6)).Length = 2)
    f = f + Chk("stdEnumerator", Not api.CoVendor_Enum().CreateFromArray(Array(1, 2)) Is Nothing)
    f = f + Chk("stdCallback", Not api.CoVendor_Callback() Is Nothing)
    If f = 0 Then
        Debug.Print "CoVendorArray 시험: 모두 통과"
        MsgBox "CoVendorArray 시험: 모두 통과", vbInformation, "칼퇴"
    Else
        MsgBox "CoVendorArray 시험: " & f & "개 실패 (직접 실행 창 확인)", vbExclamation, "칼퇴"
    End If
End Sub

Private Function Chk(ByVal testName As String, ByVal ok As Boolean) As Long
    If Not ok Then Debug.Print "[실패] " & testName: Chk = 1
End Function
