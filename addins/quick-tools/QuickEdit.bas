Attribute VB_Name = "QuickEdit"
'==============================================================
' 기능: 선택한 칸 값 고치기 4개. 값으로 고정 / IFERROR로 감싸기 / 부호 뒤집기 / 앞뒤 글자 붙이기.
' 사용법: 칸을 고르고 우클릭 > 칼퇴 빠른 도구 > 도구. 바뀔 칸 수를 알려 주고 확인받은 뒤 바꾼다.
' 설정: iferrorValue(오류일 때 보일 값, 기본 0. 빈 글자는 "" 입력), affixSkipFormula는 없음(수식 칸은 항상 건드리지 않음).
' 규칙: 값으로 고정은 수식 칸만, IFERROR는 이미 IFERROR인 칸·배열 수식은 건너뜀, 부호 뒤집기와 글자 붙이기는 상수 칸만.
' 의존: QuickTools.bas, 공용 라이브러리(CoGuard_*, CoSetting_Text)
'==============================================================
Option Explicit
Option Private Module

'--- 값으로 고정 -----------------------------------------------
Public Sub QuickEdit_ToValues()
    Dim api As Object, ws As Worksheet, areas As Collection, ar As Range, f As Variant, n As Long, tok As Long, msg As String
    If Not QT_Prepare(api, ws, areas) Then Exit Sub
    For Each ar In areas
        n = n + CountFormula(ar)
    Next ar
    If n = 0 Then api.CoGuard_Notice "수식이 있는 칸이 없습니다.": Exit Sub
    If Not QT_Confirm(api, n & "칸의 수식을 현재 값으로 바꿉니다.") Then Exit Sub
    tok = api.CoGuard_Begin(True, True)
    On Error GoTo EH
    For Each ar In areas
        If CountFormula(ar) > 0 Then ar.Value2 = ar.Value2        ' 한 번에 값으로(상수 칸은 같은 값이라 그대로)
    Next ar
    api.CoGuard_End tok
    api.CoGuard_Notice n & "칸을 값으로 바꿨습니다."
    Exit Sub
EH:
    msg = "바꾸는 중 문제가 생겼습니다: " & Err.Description
    Err.Clear
    api.CoGuard_End tok
    api.CoGuard_Notice msg
End Sub

Private Function CountFormula(ByVal ar As Range) As Long
    Dim fc As Range
    On Error Resume Next
    Set fc = ar.SpecialCells(xlCellTypeFormulas)
    If Not fc Is Nothing Then CountFormula = fc.CountLarge
End Function

'--- IFERROR로 감싸기 ---------------------------------------------
Public Sub QuickEdit_WrapIferror()
    Dim api As Object, ws As Worksheet, areas As Collection, fb As String, n As Long
    If Not QT_Prepare(api, ws, areas) Then Exit Sub
    fb = Trim$(api.CoSetting_Text(QT_ADDIN_ID, "iferrorValue", "0"))
    If Len(fb) = 0 Then fb = """"""
    n = WrapAreas(api, areas, fb, False)
    If n = 0 Then api.CoGuard_Notice "감쌀 수식이 없습니다(이미 IFERROR이거나 수식이 없음).": Exit Sub
    If Not QT_Confirm(api, n & "칸의 수식을 IFERROR(원래 수식, " & fb & ")로 감쌉니다." & vbCrLf & "(오류가 숨겨지니 값이 맞는지 따로 확인하세요.)") Then Exit Sub
    n = WrapAreas(api, areas, fb, True)
    api.CoGuard_Notice n & "칸을 감쌌습니다."
End Sub

Private Function WrapAreas(ByVal api As Object, ByVal areas As Collection, ByVal fb As String, ByVal doWrite As Boolean) As Long
    Dim ar As Range, v As Variant, i As Long, j As Long, s As String, n As Long, tok As Long, msg As String, cel As Range
    If doWrite Then tok = api.CoGuard_Begin(True, True)
    On Error GoTo EH
    For Each ar In areas
        v = QT_Read(ar)
        For i = 1 To UBound(v, 1)
            For j = 1 To UBound(v, 2)
                s = CStr(v(i, j))
                If Left$(s, 1) = "=" Then
                    If UCase$(Left$(s, 9)) <> "=IFERROR(" And Not ar.Cells(i, j).HasArray Then
                        n = n + 1
                        If doWrite Then ar.Cells(i, j).Formula = "=IFERROR(" & Mid$(s, 2) & "," & fb & ")"
                    End If
                End If
            Next j
        Next i
    Next ar
    If doWrite Then api.CoGuard_End tok
    WrapAreas = n
    Exit Function
EH:
    msg = "감싸는 중 문제가 생겼습니다: " & Err.Description
    Err.Clear
    If doWrite Then api.CoGuard_End tok
    api.CoGuard_Notice msg
    WrapAreas = n
End Function

'--- 부호 뒤집기 -----------------------------------------------
Public Sub QuickEdit_FlipSign()
    Dim api As Object, ws As Worksheet, areas As Collection, n As Long
    If Not QT_Prepare(api, ws, areas) Then Exit Sub
    n = FlipAreas(api, areas, False)
    If n = 0 Then api.CoGuard_Notice "뒤집을 숫자 칸이 없습니다.": Exit Sub
    If Not QT_Confirm(api, n & "칸의 숫자 부호를 뒤집습니다(x -1). 수식 칸·글자 칸·0은 건드리지 않습니다.") Then Exit Sub
    n = FlipAreas(api, areas, True)
    api.CoGuard_Notice n & "칸의 부호를 뒤집었습니다."
End Sub

Private Function FlipAreas(ByVal api As Object, ByVal areas As Collection, ByVal doWrite As Boolean) As Long
    Dim ar As Range, v As Variant, f As Variant, i As Long, j As Long, n As Long, tok As Long, msg As String
    If doWrite Then tok = api.CoGuard_Begin(True, True)
    On Error GoTo EH
    For Each ar In areas
        v = ar.Value2
        f = QT_Read(ar)
        If Not IsArray(v) Then
            Dim one(1 To 1, 1 To 1) As Variant
            one(1, 1) = v: v = one
        End If
        For i = 1 To UBound(v, 1)
            For j = 1 To UBound(v, 2)
                If VarType(v(i, j)) = vbDouble Then
                    If v(i, j) <> 0 And Left$(CStr(f(i, j)), 1) <> "=" Then
                        n = n + 1
                        If doWrite Then ar.Cells(i, j).Value2 = -v(i, j)
                    End If
                End If
            Next j
        Next i
    Next ar
    If doWrite Then api.CoGuard_End tok
    FlipAreas = n
    Exit Function
EH:
    msg = "바꾸는 중 문제가 생겼습니다: " & Err.Description
    Err.Clear
    If doWrite Then api.CoGuard_End tok
    api.CoGuard_Notice msg
    FlipAreas = n
End Function

'--- 앞뒤 글자 붙이기 ----------------------------------------------
Public Sub QuickEdit_AddAffix()
    Dim api As Object, ws As Worksheet, areas As Collection, pre As String, suf As String, n As Long
    If Not QT_Prepare(api, ws, areas) Then Exit Sub
    If Not api.CoGuard_AskText("각 칸 앞에 붙일 글자(없으면 비워 두세요)", pre) Then Exit Sub
    If Not api.CoGuard_AskText("각 칸 뒤에 붙일 글자(없으면 비워 두세요)", suf) Then Exit Sub
    If Len(pre) = 0 And Len(suf) = 0 Then api.CoGuard_Notice "붙일 글자가 없어 아무것도 하지 않았습니다.": Exit Sub
    n = AffixAreas(api, areas, pre, suf, False)
    If n = 0 Then api.CoGuard_Notice "값이 있는 상수 칸이 없습니다.": Exit Sub
    If Not QT_Confirm(api, n & "칸의 앞에 '" & pre & "', 뒤에 '" & suf & "'를 붙입니다(글자로 바뀝니다). 수식 칸은 건드리지 않습니다.") Then Exit Sub
    n = AffixAreas(api, areas, pre, suf, True)
    api.CoGuard_Notice n & "칸에 붙였습니다."
End Sub

Private Function AffixAreas(ByVal api As Object, ByVal areas As Collection, ByVal pre As String, ByVal suf As String, ByVal doWrite As Boolean) As Long
    Dim ar As Range, f As Variant, i As Long, j As Long, s As String, n As Long, tok As Long, msg As String, cel As Range
    If doWrite Then tok = api.CoGuard_Begin(True, True)
    On Error GoTo EH
    For Each ar In areas
        f = QT_Read(ar)
        For i = 1 To UBound(f, 1)
            For j = 1 To UBound(f, 2)
                s = CStr(f(i, j))
                If Len(s) > 0 And Left$(s, 1) <> "=" Then
                    n = n + 1
                    If doWrite Then
                        Set cel = ar.Cells(i, j)
                        s = cel.Text                                ' 보이는 글자 그대로(서식을 바꾸기 전에 읽는다)
                        If InStr(s, "##") > 0 Then s = CStr(cel.Value2)   ' 칸이 좁아 ####이면 실제 값
                        If cel.NumberFormat <> "@" Then cel.NumberFormat = "@"
                        cel.Value2 = pre & s & suf
                    End If
                End If
            Next j
        Next i
    Next ar
    If doWrite Then api.CoGuard_End tok
    AffixAreas = n
    Exit Function
EH:
    msg = "붙이는 중 문제가 생겼습니다: " & Err.Description
    Err.Clear
    If doWrite Then api.CoGuard_End tok
    api.CoGuard_Notice msg
    AffixAreas = n
End Function
