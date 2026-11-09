Attribute VB_Name = "CoVendorText"
'==============================================================
' 기능: 칼퇴 외부 라이브러리 묶음 clockout-vendor-text.xlam(글자)의 바깥 입구. 이 모듈만 Option Private Module이 없어
'       다른 통합 문서(추가기능)가 Application.Run으로 부를 수 있다. 외부 원본은 vendor.txt에 적은 파일 그대로다.
' 사용법(추가기능 쪽, 연결부 addins/_client/CoVendorClient.bas):
'   Dim v As Object
'   Set v = CoVendorClient_Api("text", "0.1.000", "내 추가기능")      ' 없거나 판이 낮으면 안내 후 Nothing
'   If v Is Nothing Then Exit Sub
'   Set rx = v.CoVendor_Regex().Create("(\d+)-(\d+)")
'   If rx.Test(s) Then s = rx.Replace(s, "$2-$1")
'   직접 부를 때: Application.Run("'clockout-vendor-text.xlam'!CoVendor_Version"), Application.Run("'clockout-vendor-text.xlam'!CoVendor_Api")
' 팩토리(개체 CoVendorTextApi의 메서드, 모두 As Object로 늦은 바인딩):
'   CoVendor_Regex() : 정규식 팩토리(stdRegex): .Create(패턴, [플래그 "i" "m" "s"]) → .Test .Match .MatchAll .Replace
'   CoVendor_Html() : HTML 팩토리(stdHTML): .Create(태그) → .CreateChild .CreateLiteral .ToString
' 주의: 판을 올릴 때 PACK_VERSION과 clockout-vendor-text.clockout.json의 version을 함께 고친다.
' 시험: Test_CoVendorText
' 의존: CoVendorTextApi.cls와 vendor.txt의 외부 원본. clockout-lib와 추가기능은 참조하지 않는다.
'==============================================================
Option Explicit

Private Const PACK_VERSION As String = "0.1.000"
Private mApi As CoVendorTextApi

Public Function CoVendor_Version() As String
    CoVendor_Version = PACK_VERSION
End Function

Public Function CoVendor_Api() As Object
    If mApi Is Nothing Then Set mApi = New CoVendorTextApi
    Set CoVendor_Api = mApi
End Function

'--- 시험 (Excel 안에서 직접 실행: 보기 > 직접 실행 창) --------------------
' 다른 통합 문서처럼 Application.Run으로 개체를 얻어 늦은 바인딩으로 부른다.
Public Sub Test_CoVendorText()
    Dim f As Long, api As Object
    Dim rx As Object, h As Object, s As String
    Set api = Application.Run("'" & ThisWorkbook.Name & "'!CoVendor_Api")
    f = f + Chk("같은 개체", api Is CoVendor_Api())
    f = f + Chk("판", Application.Run("'" & ThisWorkbook.Name & "'!CoVendor_Version") = PACK_VERSION)
    Set rx = api.CoVendor_Regex().Create("(\d+)-(\d+)")
    f = f + Chk("Test 맞음", rx.Test("번호 12-34"))
    f = f + Chk("Test 틀림", Not rx.Test("번호 없음"))
    f = f + Chk("Replace", rx.Replace("번호 12-34", "$2-$1") = "번호 34-12")
    f = f + Chk("MatchAll", api.CoVendor_Regex().Create("\d+").MatchAll("1 22 333").Count = 3)
    f = f + Chk("대소문자 무시", api.CoVendor_Regex().Create("abc", "i").Test("xABCx"))
    Set h = api.CoVendor_Html().Create("p")
    h.CreateLiteral "안녕"
    s = h.ToString()
    f = f + Chk("HTML", InStr(s, "<p") > 0 And InStr(s, "</p>") > 0)
    If f = 0 Then
        Debug.Print "CoVendorText 시험: 모두 통과"
        MsgBox "CoVendorText 시험: 모두 통과", vbInformation, "칼퇴"
    Else
        MsgBox "CoVendorText 시험: " & f & "개 실패 (직접 실행 창 확인)", vbExclamation, "칼퇴"
    End If
End Sub

Private Function Chk(ByVal testName As String, ByVal ok As Boolean) As Long
    If Not ok Then Debug.Print "[실패] " & testName: Chk = 1
End Function
