Attribute VB_Name = "ValueFix"
'==============================================================
' 기능: 값·글자 정리 도구 4개. 공백·숨은문자 정리 / 텍스트 숫자를 숫자로 / 병합 해제 + 값 채우기 / 빈 칸을 위 값으로 채우기.
' 사용법: 칸을 선택하고 우클릭 > 칼퇴 값 정리 > 도구. 바뀔 칸 수를 알려 주고 확인을 받은 뒤 바꾼다.
' 규칙: 수식 칸은 건드리지 않는다(상수만). 설정: value-tools.clockout.json의 keepLineBreak, keepLeadingZero.
' 시험: Test_ValueFix (숫자 글자 정리 규칙만, Excel 개체 없이 실행)
' 의존: ValueTools.bas(준비·확인), 공용 라이브러리(CoText_CleanInvisible, CoNum_TryParse, CoRange_HasMerged, CoGuard_*)
'==============================================================
Option Explicit
Option Private Module

'--- 공백·숨은문자 정리 -----------------------------------------
Public Sub ValueFix_CleanSpaces()
    Dim api As Object, ws As Worksheet, areas As Collection, n As Long, keepBr As Boolean
    If Not VT_Prepare(api, ws, areas) Then Exit Sub
    keepBr = api.CoSetting_Bool(VT_ADDIN_ID, "keepLineBreak", True)
    n = CleanAreas(api, areas, keepBr, False)
    If n = 0 Then VT_Done api, "정리할 칸이 없습니다.": Exit Sub
    If Not VT_Confirm(api, n & "칸의 공백·숨은 문자를 정리합니다.") Then Exit Sub
    n = CleanAreas(api, areas, keepBr, True)
    VT_Done api, n & "칸을 정리했습니다."
End Sub

Private Function CleanAreas(ByVal api As Object, ByVal areas As Collection, ByVal keepBr As Boolean, ByVal doWrite As Boolean) As Long
    Dim ar As Range, v As Variant, i As Long, j As Long, t As String, n As Long, tok As Long, msg As String
    If doWrite Then tok = api.CoGuard_Begin(True, True)
    On Error GoTo EH
    For Each ar In areas
        v = VT_Read(ar)
        For i = 1 To UBound(v, 1)
            For j = 1 To UBound(v, 2)
                If VT_IsConstText(v(i, j)) Then
                    t = api.CoText_CleanInvisible(CStr(v(i, j)), keepBr)
                    If t <> CStr(v(i, j)) Then
                        n = n + 1
                        If doWrite Then
                            If Len(t) = 0 Then ar.Cells(i, j).ClearContents Else VT_SetText ar.Cells(i, j), t
                        End If
                    End If
                End If
            Next j
        Next i
    Next ar
    If doWrite Then api.CoGuard_End tok
    CleanAreas = n
    Exit Function
EH:
    msg = "정리하는 중 문제가 생겼습니다: " & Err.Description
    Err.Clear
    If doWrite Then api.CoGuard_End tok
    api.CoGuard_Notice msg
    CleanAreas = n
End Function

'--- 텍스트 숫자를 숫자로 -----------------------------------------
Public Sub ValueFix_TextToNumber()
    Dim api As Object, ws As Worksheet, areas As Collection, n As Long, keepZero As Boolean
    If Not VT_Prepare(api, ws, areas) Then Exit Sub
    keepZero = api.CoSetting_Bool(VT_ADDIN_ID, "keepLeadingZero", True)
    n = ToNumberAreas(api, areas, keepZero, False)
    If n = 0 Then VT_Done api, "숫자로 바꿀 글자가 없습니다.": Exit Sub
    If Not VT_Confirm(api, n & "칸의 글자를 숫자로 바꿉니다." & IIf(keepZero, vbCrLf & "(0으로 시작하는 코드 모양은 바꾸지 않습니다.)", "")) Then Exit Sub
    n = ToNumberAreas(api, areas, keepZero, True)
    VT_Done api, n & "칸을 숫자로 바꿨습니다."
End Sub

Private Function ToNumberAreas(ByVal api As Object, ByVal areas As Collection, ByVal keepZero As Boolean, ByVal doWrite As Boolean) As Long
    Dim ar As Range, v As Variant, i As Long, j As Long, norm As String, x As Double, n As Long, tok As Long, msg As String
    If doWrite Then tok = api.CoGuard_Begin(True, True)
    On Error GoTo EH
    For Each ar In areas
        v = VT_Read(ar)
        For i = 1 To UBound(v, 1)
            For j = 1 To UBound(v, 2)
                If VT_IsConstText(v(i, j)) Then
                    norm = ValueFix_NormalizeNumber(api.CoText_CleanInvisible(CStr(v(i, j))))
                    If ValueFix_IsSafeNumberText(norm, keepZero) Then
                        If api.CoNum_TryParse(norm, x) Then
                            n = n + 1
                            If doWrite Then VT_SetNumber ar.Cells(i, j), x
                        End If
                    End If
                End If
            Next j
        Next i
    Next ar
    If doWrite Then api.CoGuard_End tok
    ToNumberAreas = n
    Exit Function
EH:
    msg = "숫자로 바꾸는 중 문제가 생겼습니다: " & Err.Description
    Err.Clear
    If doWrite Then api.CoGuard_End tok
    api.CoGuard_Notice msg
    ToNumberAreas = n
End Function

' 숫자 모양 글자("1,200", " 1 200 ", "123-", "(1,500)")를 "-1200" 꼴로 정리한다. 숫자 모양이 아니면 "".
Public Function ValueFix_NormalizeNumber(ByVal s As String) As String
    Dim t As String, neg As Boolean, i As Long, c As String, dots As Long, digits As Long
    t = Replace(Replace(Replace(s, ChrW(160), ""), " ", ""), ",", "")
    If Len(t) = 0 Then Exit Function
    If Left$(t, 1) = "(" And Right$(t, 1) = ")" Then neg = True: t = Mid$(t, 2, Len(t) - 2)
    If Right$(t, 1) = "-" Then
        If neg Then Exit Function
        neg = True: t = Left$(t, Len(t) - 1)
    End If
    If Left$(t, 1) = "-" Then
        If neg Then Exit Function
        neg = True: t = Mid$(t, 2)
    ElseIf Left$(t, 1) = "+" Then
        t = Mid$(t, 2)
    End If
    If Len(t) = 0 Then Exit Function
    For i = 1 To Len(t)
        c = Mid$(t, i, 1)
        If c >= "0" And c <= "9" Then
            digits = digits + 1
        ElseIf c = "." Then
            dots = dots + 1
        Else
            Exit Function
        End If
    Next i
    If digits = 0 Or dots > 1 Then Exit Function
    ValueFix_NormalizeNumber = IIf(neg, "-", "") & t
End Function

' 바꿔도 안전한 숫자 글자인지. 15자리를 넘으면 정밀도가 깨지고, keepZero면 0으로 시작하는 코드("00123")는 그대로 둔다.
Public Function ValueFix_IsSafeNumberText(ByVal norm As String, ByVal keepZero As Boolean) As Boolean
    Dim body As String, digits As Long, i As Long
    If Len(norm) = 0 Then Exit Function
    body = IIf(Left$(norm, 1) = "-", Mid$(norm, 2), norm)
    For i = 1 To Len(body)
        If Mid$(body, i, 1) <> "." Then digits = digits + 1
    Next i
    If digits > 15 Then Exit Function
    If keepZero And Len(body) > 1 Then
        If Left$(body, 1) = "0" And Mid$(body, 2, 1) <> "." Then Exit Function
    End If
    ValueFix_IsSafeNumberText = True
End Function

'--- 병합 해제 + 값 채우기 -----------------------------------------
Public Sub ValueFix_UnmergeFill()
    Dim api As Object, ws As Worksheet, areas As Collection, addrs As Collection, n As Long
    If Not VT_Prepare(api, ws, areas) Then Exit Sub
    Set addrs = FindMerges(api, areas)
    If addrs.Count = 0 Then VT_Done api, "선택한 곳에 병합된 칸이 없습니다.": Exit Sub
    If Not VT_Confirm(api, addrs.Count & "개의 병합을 풀고, 병합됐던 모든 칸에 원래 값을 채웁니다.") Then Exit Sub
    n = UnmergeAndFill(api, ws, addrs)
    VT_Done api, n & "개의 병합을 풀고 값을 채웠습니다."
End Sub

' 선택 안의 병합 영역 주소 목록(중복 없이).
Private Function FindMerges(ByVal api As Object, ByVal areas As Collection) As Collection
    Dim ar As Range, c As Range, found As New Collection, addr As String
    For Each ar In areas
        If api.CoRange_HasMerged(ar) Then
            For Each c In ar.Cells
                If c.MergeCells Then
                    If c.Address = c.MergeArea.Cells(1, 1).Address Then
                        addr = c.MergeArea.Address
                        On Error Resume Next
                        found.Add addr, addr             ' 키가 같으면 오류 -> 중복은 건너뜀
                        On Error GoTo 0
                    End If
                End If
            Next c
        End If
    Next ar
    Set FindMerges = found
End Function

Private Function UnmergeAndFill(ByVal api As Object, ByVal ws As Worksheet, ByVal addrs As Collection) As Long
    Dim k As Long, ma As Range, anchor As Variant, i As Long, tok As Long, n As Long, msg As String, cel As Range
    tok = api.CoGuard_Begin(True, True, False, True)
    On Error GoTo EH
    For k = 1 To addrs.Count
        Set ma = ws.Range(addrs(k))
        anchor = ma.Cells(1, 1).Value2
        ma.UnMerge
        For i = 2 To ma.Cells.Count                      ' 첫 칸은 원래 값(수식이면 수식)을 그대로 둔다
            Set cel = ma.Cells(i)
            If VarType(anchor) = vbString Then
                VT_SetText cel, CStr(anchor)
            ElseIf Not IsError(anchor) And Not IsEmpty(anchor) Then
                cel.Value2 = anchor
            End If
        Next i
        n = n + 1
    Next k
    api.CoGuard_End tok
    UnmergeAndFill = n
    Exit Function
EH:
    msg = "병합을 푸는 중 문제가 생겼습니다: " & Err.Description & vbCrLf & n & "개까지 처리했습니다."
    Err.Clear
    api.CoGuard_End tok
    api.CoGuard_Notice msg
    UnmergeAndFill = n
End Function

'--- 빈 칸을 위 값으로 채우기 -------------------------------------
Public Sub ValueFix_FillBlankDown()
    Dim api As Object, ws As Worksheet, areas As Collection, n As Long, skipped As Long
    If Not VT_Prepare(api, ws, areas) Then Exit Sub
    n = FillAreas(api, areas, False, skipped)
    If n = 0 Then
        VT_Done api, "채울 빈 칸이 없습니다." & IIf(skipped > 0, vbCrLf & "(병합된 칸이 있는 " & skipped & "개 덩어리는 건너뛰었습니다. 먼저 '병합 해제 + 값 채우기'를 쓰세요.)", "")
        Exit Sub
    End If
    If Not VT_Confirm(api, n & "개의 빈 칸을 바로 위 값으로 채웁니다(선택 범위 안에서, 같은 열 기준).") Then Exit Sub
    n = FillAreas(api, areas, True, skipped)
    VT_Done api, n & "개의 빈 칸을 채웠습니다." & IIf(skipped > 0, vbCrLf & "(병합된 칸이 있는 " & skipped & "개 덩어리는 건너뛰었습니다.)", "")
End Sub

Private Function FillAreas(ByVal api As Object, ByVal areas As Collection, ByVal doWrite As Boolean, ByRef skipped As Long) As Long
    Dim ar As Range, f As Variant, v As Variant, i As Long, j As Long, n As Long, tok As Long, msg As String
    Dim lastVal As Variant, hasLast As Boolean
    skipped = 0
    If doWrite Then tok = api.CoGuard_Begin(True, True)
    On Error GoTo EH
    For Each ar In areas
        If ar.Cells.Count < 2 Then
            ' 한 칸은 채울 게 없다
        ElseIf api.CoRange_HasMerged(ar) Then
            skipped = skipped + 1
        Else
            f = VT_Read(ar)
            v = VT_ReadValues(ar)
            For j = 1 To UBound(f, 2)
                hasLast = False
                For i = 1 To UBound(f, 1)
                    If Len(CStr(f(i, j))) = 0 Then
                        If hasLast Then
                            n = n + 1
                            If doWrite Then
                                If VarType(lastVal) = vbString Then VT_SetText ar.Cells(i, j), CStr(lastVal) Else ar.Cells(i, j).Value2 = lastVal
                            End If
                        End If
                    Else
                        hasLast = Not IsError(v(i, j))
                        If hasLast Then lastVal = v(i, j)
                    End If
                Next i
            Next j
        End If
    Next ar
    If doWrite Then api.CoGuard_End tok
    FillAreas = n
    Exit Function
EH:
    msg = "채우는 중 문제가 생겼습니다: " & Err.Description
    Err.Clear
    If doWrite Then api.CoGuard_End tok
    api.CoGuard_Notice msg
    FillAreas = n
End Function

'--- 시험 (직접 실행 창에서 실행) ----------------------------------
Public Sub Test_ValueFix()
    Dim f As Long
    f = f + Chk("쉼표", ValueFix_NormalizeNumber("1,200") = "1200")
    f = f + Chk("공백", ValueFix_NormalizeNumber(" 1 200 ") = "1200")
    f = f + Chk("NBSP", ValueFix_NormalizeNumber("1" & ChrW(160) & "200") = "1200")
    f = f + Chk("끝 마이너스", ValueFix_NormalizeNumber("123-") = "-123")
    f = f + Chk("괄호 음수", ValueFix_NormalizeNumber("(1,500)") = "-1500")
    f = f + Chk("앞 마이너스", ValueFix_NormalizeNumber("-7.5") = "-7.5")
    f = f + Chk("앞 플러스", ValueFix_NormalizeNumber("+5") = "5")
    f = f + Chk("소수만", ValueFix_NormalizeNumber(".5") = ".5")
    f = f + Chk("글자 섞임", ValueFix_NormalizeNumber("12원") = "")
    f = f + Chk("점 두 개", ValueFix_NormalizeNumber("1.2.3") = "")
    f = f + Chk("기호만", ValueFix_NormalizeNumber("-") = "")
    f = f + Chk("빈 글자", ValueFix_NormalizeNumber("") = "")
    f = f + Chk("앞뒤 부호 겹침", ValueFix_NormalizeNumber("-5-") = "")
    f = f + Chk("안전: 보통", ValueFix_IsSafeNumberText("1200", True))
    f = f + Chk("안전: 0 코드", Not ValueFix_IsSafeNumberText("00123", True))
    f = f + Chk("안전: 0 코드 허용", ValueFix_IsSafeNumberText("00123", False))
    f = f + Chk("안전: 0.5", ValueFix_IsSafeNumberText("0.5", True))
    f = f + Chk("안전: 그냥 0", ValueFix_IsSafeNumberText("0", True))
    f = f + Chk("안전: 16자리", Not ValueFix_IsSafeNumberText("1234567890123456", True))
    f = f + Chk("안전: 음수 0 코드", Not ValueFix_IsSafeNumberText("-007", True))
    If f = 0 Then
        Debug.Print "ValueFix 시험: 모두 통과"
        MsgBox "ValueFix 시험: 모두 통과", vbInformation, "칼퇴"
    Else
        MsgBox "ValueFix 시험: " & f & "개 실패 (직접 실행 창 확인)", vbExclamation, "칼퇴"
    End If
End Sub

Private Function Chk(ByVal testName As String, ByVal ok As Boolean) As Long
    If Not ok Then Debug.Print "[실패] " & testName: Chk = 1
End Function
