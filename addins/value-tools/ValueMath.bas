Attribute VB_Name = "ValueMath"
'==============================================================
' 기능: 숫자 도구 2개. 올림·절사·반올림(보이는 서식이 아니라 실제 값을 바꿈) / 한글 금액(일금 ○○원정).
' 사용법: 칸을 선택하고 우클릭 > 칼퇴 값 정리 > 도구. 바뀔 칸 수를 알려 주고 확인을 받은 뒤 바꾼다.
' 규칙: 값이 숫자인 상수 칸만 바꾼다(수식·글자 칸은 건드리지 않는다).
'       설정: roundHow(절사/반올림/올림), roundUnit(단위, 예 100이면 100 단위), wonPrefix, wonTarget.
' 의존: ValueTools.bas(준비·확인), 공용 라이브러리(CoNum_TruncTo, CoNum_RoundHalfUp, CoNum_ToKoreanWon, CoGuard_*)
'==============================================================
Option Explicit
Option Private Module

'--- 올림·절사·반올림 ---------------------------------------------
Public Sub ValueMath_RoundValues()
    Dim api As Object, ws As Worksheet, areas As Collection, n As Long, how As String, unitV As Double
    If Not VT_Prepare(api, ws, areas) Then Exit Sub
    how = api.CoSetting_Choice(VT_ADDIN_ID, "roundHow", "절사", "절사", "반올림", "올림")
    unitV = api.CoSetting_Number(VT_ADDIN_ID, "roundUnit", 1)
    If unitV <= 0 Then unitV = 1
    n = RoundAreas(api, areas, how, unitV, False)
    If n = 0 Then VT_Done api, "바꿀 숫자 칸이 없습니다(이미 해당 단위이거나 수식·글자 칸).": Exit Sub
    If Not VT_Confirm(api, n & "칸의 값을 " & unitV & " 단위로 " & how & "합니다." & vbCrLf & "(수식 칸은 건드리지 않습니다.)") Then Exit Sub
    n = RoundAreas(api, areas, how, unitV, True)
    VT_Done api, n & "칸을 " & unitV & " 단위로 " & how & "했습니다."
End Sub

' 한 값을 단위로 처리한 결과. 단위가 1 이상이면 정수 단위, 소수면 그 자리까지.
Public Function ValueMath_Apply(ByVal api As Object, ByVal x As Double, ByVal how As String, ByVal unitV As Double) As Double
    Dim t As Double
    t = api.CoNum_TruncTo(x, unitV)                  ' 0 쪽으로 버린 값
    Select Case how
        Case "절사"
            ValueMath_Apply = t
        Case "올림"
            If api.CoNum_Near(t, x) Then ValueMath_Apply = x Else ValueMath_Apply = t + Sgn(x) * unitV
        Case Else                                     ' 반올림: 나머지가 단위의 절반 이상이면 올림
            If api.CoNum_Near(t, x) Then
                ValueMath_Apply = x
            ElseIf Abs(x - t) * 2 >= unitV - unitV * 0.0000000001 Then
                ValueMath_Apply = t + Sgn(x) * unitV
            Else
                ValueMath_Apply = t
            End If
    End Select
    ValueMath_Apply = api.CoNum_RoundHalfUp(ValueMath_Apply, 10)   ' 부동소수 꼬리 제거
End Function

Private Function RoundAreas(ByVal api As Object, ByVal areas As Collection, ByVal how As String, _
                            ByVal unitV As Double, ByVal doWrite As Boolean) As Long
    Dim ar As Range, v As Variant, f As Variant, i As Long, j As Long, y As Double, n As Long, tok As Long, msg As String
    If doWrite Then tok = api.CoGuard_Begin(True, True)
    On Error GoTo EH
    For Each ar In areas
        v = VT_ReadValues(ar)
        f = VT_Read(ar)
        For i = 1 To UBound(v, 1)
            For j = 1 To UBound(v, 2)
                If VarType(v(i, j)) = vbDouble Then
                    If Left$(CStr(f(i, j)), 1) <> "=" Then
                        y = ValueMath_Apply(api, CDbl(v(i, j)), how, unitV)
                        If y <> CDbl(v(i, j)) Then
                            n = n + 1
                            If doWrite Then ar.Cells(i, j).Value2 = y
                        End If
                    End If
                End If
            Next j
        Next i
    Next ar
    If doWrite Then api.CoGuard_End tok
    RoundAreas = n
    Exit Function
EH:
    msg = "처리하는 중 문제가 생겼습니다: " & Err.Description
    Err.Clear
    If doWrite Then api.CoGuard_End tok
    api.CoGuard_Notice msg
    RoundAreas = n
End Function

'--- 한글 금액 -----------------------------------------------------
Public Sub ValueMath_KoreanWon()
    Dim api As Object, ws As Worksheet, areas As Collection, n As Long, pre As Boolean, same As Boolean
    If Not VT_Prepare(api, ws, areas) Then Exit Sub
    pre = api.CoSetting_Bool(VT_ADDIN_ID, "wonPrefix", True)
    same = (api.CoSetting_Choice(VT_ADDIN_ID, "wonTarget", "오른쪽 칸", "오른쪽 칸", "같은 칸") = "같은 칸")
    n = WonAreas(api, ws, areas, pre, same, False)
    If n = 0 Then VT_Done api, "한글로 바꿀 금액(0 이상 정수)이 없습니다." & IIf(same, "", vbCrLf & "(오른쪽 칸이 비어 있어야 합니다.)"): Exit Sub
    If Not VT_Confirm(api, n & "칸의 금액을 한글로 " & IIf(same, "같은 칸에 글자로 바꿉니다(숫자는 사라집니다).", "오른쪽 빈 칸에 씁니다.")) Then Exit Sub
    n = WonAreas(api, ws, areas, pre, same, True)
    VT_Done api, n & "칸을 한글 금액으로 썼습니다."
End Sub

Private Function WonAreas(ByVal api As Object, ByVal ws As Worksheet, ByVal areas As Collection, ByVal pre As Boolean, _
                          ByVal same As Boolean, ByVal doWrite As Boolean) As Long
    Dim ar As Range, v As Variant, f As Variant, i As Long, j As Long, s As String, n As Long, tok As Long, msg As String
    Dim dest As Range
    If doWrite Then tok = api.CoGuard_Begin(True, True)
    On Error GoTo EH
    For Each ar In areas
        v = VT_ReadValues(ar)
        f = VT_Read(ar)
        For i = 1 To UBound(v, 1)
            For j = 1 To UBound(v, 2)
                If VarType(v(i, j)) = vbDouble And Left$(CStr(f(i, j)), 1) <> "=" Then
                    s = api.CoNum_ToKoreanWon(CDbl(v(i, j)), pre)
                    If Len(s) > 0 Then
                        If same Then
                            Set dest = ar.Cells(i, j)
                        ElseIf ar.Cells(i, j).Column < ws.Columns.Count Then
                            Set dest = ar.Cells(i, j).Offset(0, 1)
                            If Not IsEmpty(dest.Value2) Or dest.HasFormula Then Set dest = Nothing
                        Else
                            Set dest = Nothing
                        End If
                        If Not dest Is Nothing Then
                            n = n + 1
                            If doWrite Then VT_SetText dest, s
                        End If
                    End If
                End If
            Next j
        Next i
    Next ar
    If doWrite Then api.CoGuard_End tok
    WonAreas = n
    Exit Function
EH:
    msg = "처리하는 중 문제가 생겼습니다: " & Err.Description
    Err.Clear
    If doWrite Then api.CoGuard_End tok
    api.CoGuard_Notice msg
    WonAreas = n
End Function
