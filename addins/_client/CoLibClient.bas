Attribute VB_Name = "CoLibClient"
'==============================================================
' 기능: 따로 설치하는 공용 라이브러리(clockout-lib.xlam)를 부르는 얇은 클라이언트.
'       라이브러리가 켜져 있고 판이 충분한지 확인하고, 라이브러리 개체(CoLibApi)를 돌려준다.
'       이 파일은 라이브러리를 쓰는 모든 추가기능에 똑같이 들어간다(build_addin.ps1이 자동으로 넣는다).
' 사용법(추가기능 코드에서):
'   Dim coApi As Object
'   Set coApi = CoLibClient_Api("0.1.001", "십자 음영")     ' 없거나 판이 낮으면 안내 후 Nothing
'   If coApi Is Nothing Then Exit Sub
'   coApi.CoGuard_Notice "안내 글"                          ' 늦은 바인딩: 메서드 이름은 라이브러리 함수 이름 그대로
' 규칙: 최소 판(minVersion)은 추가기능 선언(.clockout.json)의 dependencies.min_version과 같게 적는다.
'       이벤트처럼 자주 불리는 곳에서는 quiet:=True로 안내 창을 끈다(없으면 그냥 Nothing).
' 시험: Test_CoLibClient (판 비교만, Excel 개체 없이 실행)
' 의존: 없음(라이브러리가 없을 때 안내하려고 일부러 라이브러리를 쓰지 않는다)
'==============================================================
Option Explicit
Option Private Module

Private Const LIB_FILE As String = "clockout-lib.xlam"
Private Const MSG_TITLE As String = "칼퇴"

' 라이브러리 개체. 없거나 판이 낮으면 Nothing(quiet가 False면 이유를 안내 창으로 알린다).
Public Function CoLibClient_Api(ByVal minVersion As String, ByVal addinName As String, _
                                Optional ByVal quiet As Boolean = False) As Object
    Dim ver As String, why As String, api As Object
    On Error Resume Next
    ver = CStr(Application.Run("'" & LIB_FILE & "'!CoLib_Version"))
    If Err.Number <> 0 Or Len(ver) = 0 Then
        Err.Clear
        why = "'" & addinName & "'을(를) 쓰려면 공용 라이브러리(clockout-lib)가 필요합니다." & vbCrLf & _
              "칼퇴 프로그램에서 clockout-lib을 받아 설치한 뒤 Excel을 다시 켜 주세요."
    ElseIf CoLibClient_Compare(ver, minVersion) < 0 Then
        why = "'" & addinName & "'은(는) 공용 라이브러리 " & minVersion & " 이상이 필요한데 설치된 것은 " & ver & "입니다." & vbCrLf & _
              "칼퇴 프로그램에서 clockout-lib을 업데이트한 뒤 Excel을 다시 켜 주세요."
    Else
        Set api = Application.Run("'" & LIB_FILE & "'!CoLib_Api")
        If api Is Nothing Then why = "공용 라이브러리 개체를 얻지 못했습니다. Excel을 다시 켜 보세요."
    End If
    Err.Clear
    On Error GoTo 0
    If Len(why) = 0 Then
        Set CoLibClient_Api = api
    ElseIf Not quiet Then
        MsgBox why, vbExclamation, MSG_TITLE
    End If
End Function

' 판 비교("0.1.001" 꼴, 점으로 나눈 숫자를 차례로). 앞이 작으면 -1, 같으면 0, 크면 1.
Public Function CoLibClient_Compare(ByVal verA As String, ByVal verB As String) As Long
    Dim a() As String, b() As String, i As Long, n As Long, x As Long, y As Long
    a = Split(verA, "."): b = Split(verB, ".")
    n = UBound(a): If UBound(b) > n Then n = UBound(b)
    For i = 0 To n
        x = 0: y = 0
        If i <= UBound(a) Then x = CLng(Val(a(i)))
        If i <= UBound(b) Then y = CLng(Val(b(i)))
        If x < y Then CoLibClient_Compare = -1: Exit Function
        If x > y Then CoLibClient_Compare = 1: Exit Function
    Next i
End Function

'--- 시험 (직접 실행 창에서 실행) ----------------------------------
Public Sub Test_CoLibClient()
    Dim f As Long
    f = f + Chk("같음", CoLibClient_Compare("0.1.001", "0.1.001") = 0)
    f = f + Chk("작음", CoLibClient_Compare("0.1.000", "0.1.001") = -1)
    f = f + Chk("큼", CoLibClient_Compare("0.2.000", "0.1.999") = 1)
    f = f + Chk("자리수 다름", CoLibClient_Compare("1.0", "1.0.0") = 0)
    f = f + Chk("앞자리 우선", CoLibClient_Compare("1.0.0", "0.9.9") = 1)
    If f = 0 Then
        Debug.Print "CoLibClient 시험: 모두 통과"
        MsgBox "CoLibClient 시험: 모두 통과", vbInformation, MSG_TITLE
    Else
        MsgBox "CoLibClient 시험: " & f & "개 실패 (직접 실행 창 확인)", vbExclamation, MSG_TITLE
    End If
End Sub

Private Function Chk(ByVal testName As String, ByVal ok As Boolean) As Long
    If Not ok Then Debug.Print "[실패] " & testName: Chk = 1
End Function
