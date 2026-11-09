Attribute VB_Name = "QuickMisc"
'==============================================================
' 기능: 병합을 선택 영역 가운데로 바꾸기 / 텍스트 날짜를 진짜 날짜로 / 조건(병합·오류·수식·잠금 해제·빈 칸)으로 칸 선택.
' 사용법: 칸을 고르고 우클릭 > 칼퇴 빠른 도구 > 도구.
' 설정: dateFormat(날짜로 바꾼 칸의 표시 형식, 기본 yyyy-mm-dd).
' 규칙: 병합 해제는 값을 맨 위 왼쪽 칸에 두고 같은 범위를 '선택 영역의 가운데'로 정렬한다(정렬·복사가 덜 깨짐).
'       날짜는 년월일 순서만 읽는다("2026.11.08", "20261108", "26/11/08"). 못 읽는 글자는 그대로 둔다.
'       조건 선택은 값을 바꾸지 않는다. 선택이 한 칸이면 사용 중인 범위 전체에서 찾는다.
' 의존: QuickTools.bas, 공용 라이브러리(CoText_ParseDate, CoRange_*, CoGuard_*, CoSetting_Text)
'==============================================================
Option Explicit
Option Private Module

Private Const MAX_SCAN As Long = 200000

'--- 병합 -> 선택 영역 가운데 ---------------------------------------
Public Sub QuickMisc_MergeToCenter()
    Dim api As Object, ws As Worksheet, areas As Collection, ar As Range, cel As Range, list As Collection, m As Range
    Dim tok As Long, n As Long, msg As String, seen As Object, addr As Variant
    If Not QT_Prepare(api, ws, areas) Then Exit Sub
    Set list = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    For Each ar In areas
        If ar.Cells.CountLarge > MAX_SCAN Then api.CoGuard_Notice "칸이 너무 많습니다. 범위를 줄이세요.": Exit Sub
        For Each cel In ar.Cells
            If cel.MergeCells Then
                Set m = cel.MergeArea
                If Not seen.Exists(m.Address) Then seen.Add m.Address, 1: list.Add m.Address
            End If
        Next cel
    Next ar
    If list.Count = 0 Then api.CoGuard_Notice "병합된 칸이 없습니다.": Exit Sub
    If Not QT_Confirm(api, "병합 " & list.Count & "개를 풀고 '선택 영역의 가운데' 정렬로 바꿉니다.") Then Exit Sub
    tok = api.CoGuard_Begin(True, True)
    On Error GoTo EH
    For Each addr In list
        Set m = ws.Range(CStr(addr))
        m.UnMerge
        If m.Columns.Count > 1 Then
            m.HorizontalAlignment = xlCenterAcrossSelection
        End If
        n = n + 1
    Next addr
    api.CoGuard_End tok
    api.CoGuard_Notice n & "개 병합을 풀었습니다." & vbCrLf & "(여러 줄에 걸친 병합은 값이 맨 위 줄에만 남습니다.)"
    Exit Sub
EH:
    msg = "처리하는 중 문제가 생겼습니다: " & Err.Description
    Err.Clear
    api.CoGuard_End tok
    api.CoGuard_Notice msg
End Sub

'--- 텍스트 날짜 -> 날짜 ---------------------------------------------
Public Sub QuickMisc_TextToDate()
    Dim api As Object, ws As Worksheet, areas As Collection, fmt As String, n As Long
    If Not QT_Prepare(api, ws, areas) Then Exit Sub
    fmt = api.CoSetting_Text(QT_ADDIN_ID, "dateFormat", "yyyy-mm-dd")
    If Len(Trim$(fmt)) = 0 Then fmt = "yyyy-mm-dd"
    n = DateAreas(api, areas, fmt, False)
    If n = 0 Then api.CoGuard_Notice "날짜로 바꿀 글자가 없습니다.": Exit Sub
    If Not QT_Confirm(api, n & "칸의 글자를 날짜로 바꿉니다(표시 형식 " & fmt & ").") Then Exit Sub
    n = DateAreas(api, areas, fmt, True)
    api.CoGuard_Notice n & "칸을 날짜로 바꿨습니다."
End Sub

Private Function DateAreas(ByVal api As Object, ByVal areas As Collection, ByVal fmt As String, ByVal doWrite As Boolean) As Long
    Dim ar As Range, v As Variant, f As Variant, i As Long, j As Long, s As String, d As Date, n As Long, tok As Long, msg As String
    If doWrite Then tok = api.CoGuard_Begin(True, True)
    On Error GoTo EH
    For Each ar In areas
        f = QT_Read(ar)
        v = ar.Value2
        If Not IsArray(v) Then
            Dim one(1 To 1, 1 To 1) As Variant
            one(1, 1) = v: v = one
        End If
        For i = 1 To UBound(f, 1)
            For j = 1 To UBound(f, 2)
                If Left$(CStr(f(i, j)), 1) <> "=" Then
                    s = ""
                    If VarType(v(i, j)) = vbString Then
                        s = Trim$(CStr(v(i, j)))
                    ElseIf VarType(v(i, j)) = vbDouble Then
                        If v(i, j) >= 19000101# And v(i, j) <= 20991231# And v(i, j) = Fix(v(i, j)) Then s = CStr(CLng(v(i, j)))   ' 20261108 같은 숫자
                    End If
                    If Len(s) > 0 Then
                        If api.CoText_ParseDate(s, d) Then
                            n = n + 1
                            If doWrite Then
                                ar.Cells(i, j).NumberFormat = fmt
                                ar.Cells(i, j).Value2 = CDbl(d)
                                If ar.Cells(i, j).Text Like "#*" And InStr(ar.Cells(i, j).Text, "##") > 0 Then ar.Columns(j).EntireColumn.AutoFit   ' 좁아서 ####이면 열을 넓힌다
                            End If
                        End If
                    End If
                End If
            Next j
        Next i
    Next ar
    If doWrite Then api.CoGuard_End tok
    DateAreas = n
    Exit Function
EH:
    msg = "바꾸는 중 문제가 생겼습니다: " & Err.Description
    Err.Clear
    If doWrite Then api.CoGuard_End tok
    api.CoGuard_Notice msg
    DateAreas = n
End Function

'--- 조건으로 칸 선택 -------------------------------------------------
Public Sub QuickMisc_SelectSpecial()
    Dim api As Object, ws As Worksheet, why As String, sel As Range, scope As Range, pick As String, hit As Range
    Dim cel As Range, cnt As Long
    Set api = QT_Api()
    If api Is Nothing Then Exit Sub
    If Not api.CoRange_GetSelection(ws, sel, why, True, True) Then api.CoGuard_Notice why: Exit Sub
    If sel.Cells.CountLarge = 1 Then Set scope = ws.UsedRange Else Set scope = Intersect(sel, ws.UsedRange)
    If scope Is Nothing Then api.CoGuard_Notice "선택한 칸에 값이 없습니다.": Exit Sub
    If Not api.CoGuard_AskText("어떤 칸을 고를까요? 번호를 입력하세요." & vbCrLf & "1 병합된 칸   2 오류 칸   3 수식 칸   4 잠금 해제된 칸   5 빈 칸   6 숨겨진 칸 제외한 보이는 칸", pick, , "1") Then Exit Sub
    On Error Resume Next
    Select Case Trim$(pick)
        Case "1"
            If scope.Cells.CountLarge > MAX_SCAN Then api.CoGuard_Notice "범위가 너무 큽니다.": Exit Sub
            For Each cel In scope.Cells
                If cel.MergeCells Then
                    cnt = cnt + 1
                    If hit Is Nothing Then Set hit = cel.MergeArea Else Set hit = Union(hit, cel.MergeArea)
                    If cnt >= 5000 Then Exit For
                End If
            Next cel
        Case "2": Set hit = scope.SpecialCells(xlCellTypeFormulas, xlErrors)
                  If hit Is Nothing Then Set hit = scope.SpecialCells(xlCellTypeConstants, xlErrors)
        Case "3": Set hit = scope.SpecialCells(xlCellTypeFormulas)
        Case "4"
            If scope.Cells.CountLarge > MAX_SCAN Then api.CoGuard_Notice "범위가 너무 큽니다.": Exit Sub
            For Each cel In scope.Cells
                If Not cel.Locked Then
                    cnt = cnt + 1
                    If hit Is Nothing Then Set hit = cel Else Set hit = Union(hit, cel)
                    If cnt >= 5000 Then Exit For
                End If
            Next cel
        Case "5": Set hit = scope.SpecialCells(xlCellTypeBlanks)
                  If Not hit Is Nothing Then Set hit = DropMerged(hit)         ' 병합 칸은 빈 칸으로 세지 않는다
        Case "6": Set hit = scope.SpecialCells(xlCellTypeVisible)
        Case Else
            api.CoGuard_Notice "1~6 중 하나를 입력하세요.": Exit Sub
    End Select
    On Error GoTo 0
    If hit Is Nothing Then api.CoGuard_Notice "조건에 맞는 칸이 없습니다.": Exit Sub
    hit.Select
    Application.StatusBar = "조건에 맞는 칸 " & hit.CountLarge & "칸을 골랐습니다."
End Sub

' 범위에서 병합된 칸을 뺀 나머지(없으면 Nothing). 칸이 너무 많으면 그대로 돌려준다.
Private Function DropMerged(ByVal rng As Range) As Range
    Dim cel As Range, keep As Range, cnt As Long
    If rng.CountLarge > MAX_SCAN Then Set DropMerged = rng: Exit Function
    For Each cel In rng.Cells
        If Not cel.MergeCells Then
            cnt = cnt + 1
            If keep Is Nothing Then Set keep = cel Else Set keep = Union(keep, cel)
            If cnt >= 5000 Then Exit For
        End If
    Next cel
    Set DropMerged = keep
End Function
