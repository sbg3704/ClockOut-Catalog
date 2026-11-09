Attribute VB_Name = "CoVendorData"
'==============================================================
' 기능: 칼퇴 외부 라이브러리 묶음 clockout-vendor-data.xlam(형식)의 바깥 입구. 이 모듈만 Option Private Module이 없어
'       다른 통합 문서(추가기능)가 Application.Run으로 부를 수 있다. 외부 원본은 vendor.txt에 적은 파일 그대로다.
' 사용법(추가기능 쪽, 연결부 addins/_client/CoVendorClient.bas):
'   Dim v As Object
'   Set v = CoVendorClient_Api("data", "0.1.000", "내 추가기능")      ' 없거나 판이 낮으면 안내 후 Nothing
'   If v Is Nothing Then Exit Sub
'   Set j = v.CoVendor_Json()
'   Set o = j.Parse("{""a"":[1,2]}")      ' 개체는 Dictionary, 배열은 Collection
'   s = j.Stringify(o, 2)                    ' 들여쓰기 2칸
'   직접 부를 때: Application.Run("'clockout-vendor-data.xlam'!CoVendor_Version"), Application.Run("'clockout-vendor-data.xlam'!CoVendor_Api")
' 팩토리(개체 CoVendorDataApi의 메서드, 모두 As Object로 늦은 바인딩):
'   CoVendor_Json() : JSON·날짜 도구(CoVendorJson): .Parse(글) .Stringify(값, [들여쓰기]) .ParseIso .ToIso .ToUtc .FromUtc
'   CoVendor_Dict() : 빈 사전(VBA-Dictionary): .Add 키, 값 / .Exists(키) / d(키)
' 주의: 판을 올릴 때 PACK_VERSION과 clockout-vendor-data.clockout.json의 version을 함께 고친다.
' 시험: Test_CoVendorData
' 의존: CoVendorDataApi.cls와 vendor.txt의 외부 원본. clockout-lib와 추가기능은 참조하지 않는다.
'==============================================================
Option Explicit

Private Const PACK_VERSION As String = "0.1.000"
Private mApi As CoVendorDataApi

Public Function CoVendor_Version() As String
    CoVendor_Version = PACK_VERSION
End Function

Public Function CoVendor_Api() As Object
    If mApi Is Nothing Then Set mApi = New CoVendorDataApi
    Set CoVendor_Api = mApi
End Function

'--- 시험 (Excel 안에서 직접 실행: 보기 > 직접 실행 창) --------------------
' 다른 통합 문서처럼 Application.Run으로 개체를 얻어 늦은 바인딩으로 부른다.
Public Sub Test_CoVendorData()
    Dim f As Long, api As Object
    Dim j As Object, o As Object, o2 As Object, d As Object, s As String, dt As Date
    Set api = Application.Run("'" & ThisWorkbook.Name & "'!CoVendor_Api")
    f = f + Chk("같은 개체", api Is CoVendor_Api())
    f = f + Chk("판", Application.Run("'" & ThisWorkbook.Name & "'!CoVendor_Version") = PACK_VERSION)
    Set j = api.CoVendor_Json()
    Set o = j.Parse("{""name"":""칼퇴"",""n"":12.5,""ok"":true,""items"":[1,2,{""c"":""한글""}],""none"":null}")
    f = f + Chk("JSON 개체", Not o Is Nothing)
    If Not o Is Nothing Then
        f = f + Chk("글자", o("name") = "칼퇴")
        f = f + Chk("숫자", o("n") = 12.5)
        f = f + Chk("참거짓", o("ok") = True)
        f = f + Chk("배열 개수", o("items").Count = 3)
        f = f + Chk("중첩", o("items")(3)("c") = "한글")
        f = f + Chk("null", IsNull(o("none")))
        o("name") = "칼퇴 예제"
        o("items").Add 4
        s = j.Stringify(o)
        Set o2 = j.Parse(s)
        f = f + Chk("되읽기 글자", o2("name") = "칼퇴 예제")
        f = f + Chk("되읽기 배열", o2("items").Count = 4)
        f = f + Chk("들여쓰기", InStr(j.Stringify(o, 2), vbCrLf) > 0)
    End If
    f = f + Chk("잘못된 JSON은 오류", BadJson(j))
    Set d = api.CoVendor_Dict()
    d.Add "x", 1
    f = f + Chk("사전", d.Exists("x") And d("x") = 1 And d.Count = 1)
    dt = DateSerial(2026, 11, 9) + TimeSerial(15, 6, 0)
    s = j.ToIso(dt)
    f = f + Chk("ISO 모양", Len(s) >= 20 And Right$(s, 1) = "Z")
    f = f + Chk("ISO 되읽기", Abs(j.ParseIso(s) - dt) < 1 / 86400)
    f = f + Chk("UTC 왕복", Abs(j.FromUtc(j.ToUtc(dt)) - dt) < 1 / 86400)
    If f = 0 Then
        Debug.Print "CoVendorData 시험: 모두 통과"
        MsgBox "CoVendorData 시험: 모두 통과", vbInformation, "칼퇴"
    Else
        MsgBox "CoVendorData 시험: " & f & "개 실패 (직접 실행 창 확인)", vbExclamation, "칼퇴"
    End If
End Sub

Private Function Chk(ByVal testName As String, ByVal ok As Boolean) As Long
    If Not ok Then Debug.Print "[실패] " & testName: Chk = 1
End Function

' 잘못된 JSON을 주면 오류가 나는지(원본 동작: 오류를 던진다).
Private Function BadJson(ByVal j As Object) As Boolean
    Dim o As Object
    On Error Resume Next
    Set o = j.Parse("{""a"":")
    BadJson = (Err.Number <> 0)
    Err.Clear
End Function
