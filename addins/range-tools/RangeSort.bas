Attribute VB_Name = "RangeSort"
'==============================================================
' 기능: 엑셀 기본 정렬은 "1호, 10호, 2호"로 줄 세운다. 숫자를 크기로 읽어 "1호, 2호, 10호" 순으로 줄 단위 정렬한다.
' 사용법: 정렬할 표를 고르고(한 칸만 누르면 이어진 표 전체) 우클릭 > 칼퇴 범위 도구 > 자연 정렬.
'         기준은 현재 활성 칸이 있는 열(선택 밖이면 첫 열). 서식·값·수식이 줄과 함께 움직인다(엑셀 정렬과 같음).
' 설정: sortHeader(첫 줄은 제목이라 정렬 제외, 기본 켬), sortDescending(내림차순, 기본 끔).
' 규칙: 병합 칸이 있으면 하지 않는다. 최대 5000줄. 임시 시트를 잠깐 만들었다 지운다(원본 시트 구조는 그대로).
' 의존: RangeTools.bas, 공용 라이브러리(CoText_NaturalCompare, CoSetting_Bool, CoRange_HasMerged, CoGuard_*)
'==============================================================
Option Explicit
Option Private Module

Private Const MAX_ROWS As Long = 5000

Public Sub RangeSort_Natural()
    Dim api As Object, ws As Worksheet, rng As Range, body As Range, hasHdr As Boolean, desc As Boolean
    Dim keyCol As Long, v As Variant, n As Long, i As Long, perm() As Long, keys() As String, tok As Long, msg As String
    Dim tmp As Worksheet, wb As Workbook
    Set api = RT_Api()
    If api Is Nothing Then Exit Sub
    Set rng = RT_Area(api, ws)
    If rng Is Nothing Then Exit Sub
    If rng.Areas.Count > 1 Then api.CoGuard_Notice "여러 덩어리가 아니라 한 덩어리를 고르세요.": Exit Sub
    If api.CoRange_HasMerged(rng) Then api.CoGuard_Notice "병합된 칸이 있어 정렬할 수 없습니다. 병합을 먼저 푸세요.": Exit Sub
    If Not api.CoRange_CanEdit(rng) Then api.CoGuard_Notice "시트가 보호되어 있고 잠긴 칸이 있어 정렬할 수 없습니다.": Exit Sub
    hasHdr = api.CoSetting_Bool(RT_ADDIN_ID, "sortHeader", True)
    desc = api.CoSetting_Bool(RT_ADDIN_ID, "sortDescending", False)
    If hasHdr Then
        If rng.Rows.Count < 3 Then api.CoGuard_Notice "정렬할 줄이 부족합니다.": Exit Sub
        Set body = rng.Offset(1, 0).Resize(rng.Rows.Count - 1)
    Else
        If rng.Rows.Count < 2 Then api.CoGuard_Notice "정렬할 줄이 부족합니다.": Exit Sub
        Set body = rng
    End If
    n = body.Rows.Count
    If n > MAX_ROWS Then api.CoGuard_Notice "줄이 너무 많습니다(" & n & "줄, 최대 " & MAX_ROWS & "줄)." : Exit Sub
    keyCol = 1
    If Not Intersect(ActiveCell, rng) Is Nothing Then keyCol = ActiveCell.Column - rng.Column + 1
    v = api.CoRange_Values2D(body.Columns(keyCol))
    ReDim keys(1 To n): ReDim perm(1 To n)
    For i = 1 To n
        If IsError(v(i, 1)) Then keys(i) = "" Else keys(i) = CStr(v(i, 1))
        perm(i) = i
    Next i
    MergeSortIdx api, keys, perm, 1, n, desc
    If Not RT_Confirm(api, n & "줄을 '" & ws.Cells(rng.Row, body.Columns(keyCol).Column).Address(False, False) & "' 열 기준 자연 순서로 " & _
                      IIf(desc, "내림차순", "오름차순") & " 정렬합니다." & IIf(hasHdr, vbCrLf & "(첫 줄은 제목으로 두고 제외)", "")) Then Exit Sub
    Set wb = ws.Parent
    tok = api.CoGuard_Begin(True, True, False, True)
    On Error GoTo EH
    Set tmp = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    For i = 1 To n
        body.Rows(perm(i)).Copy tmp.Cells(i, 1)
    Next i
    tmp.Range(tmp.Cells(1, 1), tmp.Cells(n, body.Columns.Count)).Copy
    body.PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    tmp.Delete
    ws.Activate
    body.Cells(1, 1).Select
    api.CoGuard_End tok
    api.CoGuard_Notice n & "줄을 자연 순서로 정렬했습니다."
    Exit Sub
EH:
    msg = "정렬하는 중 문제가 생겼습니다: " & Err.Description & vbCrLf & "(표가 일부만 바뀌었을 수 있으니 저장하지 말고 확인하세요.)"
    Err.Clear
    On Error Resume Next
    Application.CutCopyMode = False
    tmp.Delete
    ws.Activate
    api.CoGuard_End tok
    api.CoGuard_Notice msg
End Sub

' 자연 순서 병합 정렬(안정 정렬). perm에 정렬된 원래 위치 번호가 담긴다.
Private Sub MergeSortIdx(ByVal api As Object, ByRef keys() As String, ByRef perm() As Long, ByVal lo As Long, ByVal hi As Long, ByVal desc As Boolean)
    Dim tmp() As Long, i As Long, j As Long, k As Long, mid1 As Long, c As Long
    If hi <= lo Then Exit Sub
    mid1 = (lo + hi) \ 2
    MergeSortIdx api, keys, perm, lo, mid1, desc
    MergeSortIdx api, keys, perm, mid1 + 1, hi, desc
    ReDim tmp(lo To hi)
    i = lo: j = mid1 + 1: k = lo
    Do While i <= mid1 And j <= hi
        c = api.CoText_NaturalCompare(keys(perm(i)), keys(perm(j)))
        If desc Then c = -c
        If c <= 0 Then tmp(k) = perm(i): i = i + 1 Else tmp(k) = perm(j): j = j + 1
        k = k + 1
    Loop
    Do While i <= mid1: tmp(k) = perm(i): i = i + 1: k = k + 1: Loop
    Do While j <= hi: tmp(k) = perm(j): j = j + 1: k = k + 1: Loop
    For k = lo To hi: perm(k) = tmp(k): Next k
End Sub
