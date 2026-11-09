Attribute VB_Name = "CoVendorClient"
'==============================================================
' 기능: 따로 설치하는 외부 라이브러리 묶음(clockout-vendor-data/text/array/sys.xlam)을 부르는 얇은 연결부.
'       묶음이 켜져 있고 판이 충분한지 확인하고, 묶음의 진입점 개체를 돌려준다.
'       선언 dependencies에 clockout-vendor-* 가 있으면 tools/build_addin.ps1이 이 파일을 자동으로 넣는다.
' 사용법(추가기능 코드에서):
'   Dim v As Object, j As Object
'   Set v = CoVendorClient_Api("data", "0.1.000", "JSON 예제")    ' 묶음: data·text·array·sys (또는 전체 id)
'   If v Is Nothing Then Exit Sub                                  ' 없거나 판이 낮으면 안내 후 Nothing
'   Set j = v.CoVendor_Json()                                      ' 메서드 목록은 각 묶음의 README.md
' 규칙: 최소 판(minVersion)은 선언의 dependencies.min_version과 같게 적는다. 자주 불리는 곳은 quiet:=True.
'       한 번 얻은 개체는 모듈 변수에 두고 다시 써도 된다(묶음은 같은 개체를 계속 돌려준다).
' 시험: Test_CoVendorClient (묶음 이름·판 비교만, Excel 개체 없이 실행)
' 의존: 없음(묶음이 없을 때 안내하려고 일부러 아무것도 쓰지 않는다. CoLibClient와 함께 넣어도 이름이 겹치지 않는다)
'==============================================================
Option Explicit
Option Private Module

Private Const PACK_PREFIX As String = "clockout-vendor-"
Private Const MSG_TITLE As String = "칼퇴"

' 묶음 진입점 개체. 없거나 판이 낮으면 Nothing(quiet가 False면 이유를 안내 창으로 알린다).
Public Function CoVendorClient_Api(ByVal pack As String, ByVal minVersion As String, ByVal addinName As String, _
                                   Optional ByVal quiet As Boolean = False) As Object
    Dim packId As String, ver As String, why As String, api As Object
    packId = CoVendorClient_PackId(pack)
    If Len(packId) = 0 Then
        why = "'" & addinName & "': 외부 라이브러리 묶음 이름이 올바르지 않습니다(" & pack & ")."
    Else
        On Error Resume Next
        ver = CStr(Application.Run("'" & packId & ".xlam'!CoVendor_Version"))
        If Err.Number <> 0 Or Len(ver) = 0 Then
            why = "'" & addinName & "'을(를) 쓰려면 외부 라이브러리(" & packId & ")가 필요합니다." & vbCrLf & _
                  "칼퇴 프로그램에서 " & packId & "을(를) 받아 설치한 뒤 Excel을 다시 켜 주세요."
        ElseIf VerCompare(ver, minVersion) < 0 Then
            why = "'" & addinName & "'은(는) " & packId & " " & minVersion & " 이상이 필요한데 설치된 것은 " & ver & "입니다." & vbCrLf & _
                  "칼퇴 프로그램에서 " & packId & "을(를) 업데이트한 뒤 Excel을 다시 켜 주세요."
        Else
            Set api = Application.Run("'" & packId & ".xlam'!CoVendor_Api")
            If api Is Nothing Then why = packId & " 개체를 얻지 못했습니다. Excel을 다시 켜 보세요."
        End If
        Err.Clear
        On Error GoTo 0
    End If
    If Len(why) = 0 Then
        Set CoVendorClient_Api = api
    ElseIf Not quiet Then
        MsgBox why, vbExclamation, MSG_TITLE
    End If
End Function

' "data" 또는 "clockout-vendor-data" -> "clockout-vendor-data". 영문 소문자·숫자·하이픈이 아니면 "".
Public Function CoVendorClient_PackId(ByVal pack As String) As String
    Dim s As String, i As Long, ch As String
    s = LCase$(Trim$(pack))
    If Left$(s, Len(PACK_PREFIX)) <> PACK_PREFIX Then s = PACK_PREFIX & s
    If Len(s) = Len(PACK_PREFIX) Then Exit Function
    For i = Len(PACK_PREFIX) + 1 To Len(s)
        ch = Mid$(s, i, 1)
        If Not (ch Like "[a-z0-9-]") Then Exit Function
    Next i
    CoVendorClient_PackId = s
End Function

' 판 비교("0.1.000" 꼴). 앞이 작으면 -1, 같으면 0, 크면 1.
Private Function VerCompare(ByVal verA As String, ByVal verB As String) As Long
    Dim a() As String, b() As String, i As Long, n As Long, x As Long, y As Long
    a = Split(verA, "."): b = Split(verB, ".")
    n = UBound(a): If UBound(b) > n Then n = UBound(b)
    For i = 0 To n
        x = 0: y = 0
        If i <= UBound(a) Then x = CLng(Val(a(i)))
        If i <= UBound(b) Then y = CLng(Val(b(i)))
        If x < y Then VerCompare = -1: Exit Function
        If x > y Then VerCompare = 1: Exit Function
    Next i
End Function

'--- 시험 (직접 실행 창에서 실행) ----------------------------------
Public Sub Test_CoVendorClient()
    Dim f As Long
    f = f + Chk("짧은 이름", CoVendorClient_PackId("data") = "clockout-vendor-data")
    f = f + Chk("전체 id", CoVendorClient_PackId("clockout-vendor-sys") = "clockout-vendor-sys")
    f = f + Chk("대문자·공백", CoVendorClient_PackId(" Text ") = "clockout-vendor-text")
    f = f + Chk("빈 이름", CoVendorClient_PackId("") = "")
    f = f + Chk("이상한 글자", CoVendorClient_PackId("da'ta") = "")
    f = f + Chk("판 같음", VerCompare("0.1.000", "0.1.000") = 0)
    f = f + Chk("판 작음", VerCompare("0.1.000", "0.1.001") = -1)
    f = f + Chk("판 큼", VerCompare("0.2.000", "0.1.999") = 1)
    f = f + Chk("없는 묶음은 Nothing", CoVendorClient_Api("zz-none", "0.1.000", "시험", True) Is Nothing)
    If f = 0 Then
        Debug.Print "CoVendorClient 시험: 모두 통과"
        MsgBox "CoVendorClient 시험: 모두 통과", vbInformation, MSG_TITLE
    Else
        MsgBox "CoVendorClient 시험: " & f & "개 실패 (직접 실행 창 확인)", vbExclamation, MSG_TITLE
    End If
End Sub

Private Function Chk(ByVal testName As String, ByVal ok As Boolean) As Long
    If Not ok Then Debug.Print "[실패] " & testName: Chk = 1
End Function
