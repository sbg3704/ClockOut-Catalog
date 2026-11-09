Attribute VB_Name = "VendorSample"
'@clockout id=vendor-sample name=JSON 보기 좋게(예제) host=excel
'@desc 선택한 칸의 JSON 글을 보기 좋게(2칸 들여쓰기) 바꾸거나 한 줄로 줄인다. 외부 라이브러리 묶음 사용 예제.
'==============================================================
' 기능: 외부 라이브러리 묶음 clockout-vendor-data(VBA-JSON)를 CoVendorClient로 불러 JSON을 읽고(Parse) 쓴다(Stringify).
'       소스를 복사해 넣지 않는다. 이 추가기능의 .xlam에는 이 모듈과 연결부(CoLibClient·CoVendorClient)만 들어간다.
' 사용법: 리본 칼퇴 탭 > "JSON 보기 좋게" / "JSON 한 줄로". JSON이 아닌 칸·수식 칸은 건너뛰고 몇 칸 바꿨는지 알려 준다.
' 의존: clockout-lib 0.1.001 이상(CoRange_GetSelection, CoGuard_Notice), clockout-vendor-data 0.1.000 이상(CoVendor_Json)
' 시험: Test_VendorSample (묶음을 불러 JSON 읽기·고치기·쓰기·되읽기. 빌드 시험에서는 out\의 묶음 .xlam을 열어 둔다)
'==============================================================
Option Explicit
Option Private Module

Private Const ADDIN_NAME As String = "JSON 보기 좋게(예제)"
Private Const MIN_LIB As String = "0.1.001"
Private Const MIN_DATA As String = "0.1.000"
Private Const MAX_CELLS As Long = 10000

Public Sub VendorSample_Pretty()
    RunFormat 2
End Sub

Public Sub VendorSample_Compact()
    RunFormat 0
End Sub

' indent: 0이면 한 줄, 아니면 그 칸 수만큼 들여쓰기.
Private Sub RunFormat(ByVal indent As Long)
    Dim lib As Object, j As Object, ws As Worksheet, rng As Range, c As Range
    Dim why As String, outText As String, n As Long, skipped As Long
    Set lib = CoLibClient_Api(MIN_LIB, ADDIN_NAME): If lib Is Nothing Then Exit Sub
    Set j = VendorSample_Json(False): If j Is Nothing Then Exit Sub
    If Not lib.CoRange_GetSelection(ws, rng, why) Then lib.CoGuard_Notice why: Exit Sub
    If rng.CountLarge > MAX_CELLS Then lib.CoGuard_Notice "칸이 너무 많습니다(최대 " & MAX_CELLS & "칸).": Exit Sub
    For Each c In rng.Cells
        If Not c.HasFormula And VarType(c.Value) = vbString Then
            If VendorSample_Reformat(j, CStr(c.Value), indent, outText) Then
                c.Value = outText
                n = n + 1
            Else
                skipped = skipped + 1
            End If
        End If
    Next c
    lib.CoGuard_Notice n & "칸을 바꿨습니다." & IIf(skipped > 0, " (JSON이 아니라 건너뛴 칸 " & skipped & "개)", "")
End Sub

' 묶음의 JSON 도구. 없으면 Nothing(quiet가 아니면 안내).
Public Function VendorSample_Json(ByVal quiet As Boolean) As Object
    Dim v As Object
    Set v = CoVendorClient_Api("data", MIN_DATA, ADDIN_NAME, quiet)
    If Not v Is Nothing Then Set VendorSample_Json = v.CoVendor_Json()
End Function

' JSON 글을 읽어 다시 쓴다. 읽을 수 없으면 False(오류를 밖으로 던지지 않는다).
Public Function VendorSample_Reformat(ByVal j As Object, ByVal jsonText As String, ByVal indent As Long, _
                                      ByRef outText As String) As Boolean
    Dim o As Object
    outText = ""
    If Len(Trim$(jsonText)) = 0 Then Exit Function
    On Error GoTo Fail
    Set o = j.Parse(jsonText)
    If indent > 0 Then outText = j.Stringify(o, indent) Else outText = j.Stringify(o)
    VendorSample_Reformat = True
    Exit Function
Fail:
    outText = ""
End Function

'--- 시험 (직접 실행 창에서 실행. clockout-vendor-data.xlam이 열려 있어야 한다) ---------
Public Sub Test_VendorSample()
    Dim f As Long, j As Object, o As Object, s As String
    Set j = VendorSample_Json(True)
    f = f + Chk("묶음 연결", Not j Is Nothing)
    If Not j Is Nothing Then
        Set o = j.Parse("{""이름"":""칼퇴"",""목록"":[1,2,3],""켜짐"":true}")
        o("이름") = o("이름") & " 예제"
        o("목록").Add 4
        s = j.Stringify(o)
        f = f + Chk("쓰기", InStr(s, """목록"":[1,2,3,4]") > 0 Or InStr(s, "[1,2,3,4]") > 0)
        Set o = j.Parse(s)
        f = f + Chk("되읽기 글자", o("이름") = "칼퇴 예제")
        f = f + Chk("되읽기 배열", o("목록").Count = 4 And o("목록")(4) = 4)
        f = f + Chk("보기 좋게", VendorSample_Reformat(j, "{""a"":[1,2]}", 2, s) And InStr(s, vbCrLf) > 0)
        f = f + Chk("한 줄로", VendorSample_Reformat(j, "{" & vbCrLf & "  ""a"": 1" & vbCrLf & "}", 0, s) And s = "{""a"":1}")
        f = f + Chk("JSON 아님", Not VendorSample_Reformat(j, "그냥 글", 2, s) And s = "")
    End If
    If f = 0 Then
        Debug.Print "VendorSample 시험: 모두 통과"
        MsgBox "VendorSample 시험: 모두 통과", vbInformation, "칼퇴"
    Else
        MsgBox "VendorSample 시험: " & f & "개 실패 (직접 실행 창 확인)", vbExclamation, "칼퇴"
    End If
End Sub

Private Function Chk(ByVal testName As String, ByVal ok As Boolean) As Long
    If Not ok Then Debug.Print "[실패] " & testName: Chk = 1
End Function
