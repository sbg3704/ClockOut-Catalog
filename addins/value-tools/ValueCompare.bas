Attribute VB_Name = "ValueCompare"
'==============================================================
' 기능: 두 열을 줄 단위로 비교해 서로 다른 줄의 칸을 골라 준다(색칠·수식·값 변경 없음, 선택만 바꿈).
' 사용법: (1) 나란한 두 열을 한 덩어리로 드래그하거나 (2) Ctrl을 누른 채 두 열을 따로 골라 우클릭 > 칼퇴 값 정리 > 두 열 다른 칸 선택.
'         같은 위치(첫 칸부터 같은 줄 수)끼리 비교한다. 앞뒤 공백은 같은 것으로 본다.
' 설정: compareIgnoreCase(대소문자 무시, 기본 켬). 골라 주는 칸은 최대 3000줄(너무 많으면 나머지는 개수만 알림).
' 의존: ValueTools.bas(준비), 공용 라이브러리(CoSetting_Bool, CoRange_*, CoGuard_Notice)
'==============================================================
Option Explicit
Option Private Module

Private Const MAX_ROWS As Long = 3000

Public Sub ValueCompare_SelectDiff()
    Dim api As Object, ws As Worksheet, areas As Collection, a As Range, b As Range
    Dim va As Variant, vb As Variant, i As Long, nRows As Long, diff As Long, shown As Long
    Dim ignoreCase As Boolean, pick As Range, sa As String, sb As String, same As Boolean
    If Not VT_Prepare(api, ws, areas, False) Then Exit Sub
    ignoreCase = api.CoSetting_Bool(VT_ADDIN_ID, "compareIgnoreCase", True)
    If Not PickTwo(api, ws, a, b) Then Exit Sub
    nRows = a.Rows.Count
    va = VT_ReadValues(a)
    vb = VT_ReadValues(b)
    For i = 1 To nRows
        sa = Trim$(CellText(va(i, 1)))
        sb = Trim$(CellText(vb(i, 1)))
        If ignoreCase Then same = (StrComp(sa, sb, vbTextCompare) = 0) Else same = (StrComp(sa, sb, vbBinaryCompare) = 0)
        If Not same Then
            diff = diff + 1
            If shown < MAX_ROWS Then
                shown = shown + 1
                If pick Is Nothing Then Set pick = Union(a.Cells(i, 1), b.Cells(i, 1)) Else Set pick = Union(pick, a.Cells(i, 1), b.Cells(i, 1))
            End If
        End If
    Next i
    If diff = 0 Then VT_Done api, "두 열이 " & nRows & "줄 모두 같습니다.": Exit Sub
    pick.Select
    If diff > shown Then
        VT_Done api, "다른 줄 " & diff & "개 중 앞의 " & shown & "줄만 골랐습니다."
    Else
        Application.StatusBar = "다른 줄 " & diff & "개를 골랐습니다 (전체 " & nRows & "줄)"
    End If
End Sub

' 비교할 두 열 범위를 정한다. 안 되면 안내 후 False.
Private Function PickTwo(ByVal api As Object, ByVal ws As Worksheet, ByRef a As Range, ByRef b As Range) As Boolean
    Dim sel As Range, why As String, r1 As Range, r2 As Range, lastRow As Long, n As Long
    If Not api.CoRange_GetSelection(ws, sel, why, True, True) Then api.CoGuard_Notice why: Exit Function
    If sel.Areas.Count = 1 And sel.Columns.Count = 2 Then
        Set r1 = sel.Columns(1): Set r2 = sel.Columns(2)
    ElseIf sel.Areas.Count = 2 And sel.Areas(1).Columns.Count = 1 And sel.Areas(2).Columns.Count = 1 Then
        Set r1 = sel.Areas(1): Set r2 = sel.Areas(2)
        If r1.Rows.Count <> r2.Rows.Count Then
            api.CoGuard_Notice "두 열의 줄 수가 다릅니다(" & r1.Rows.Count & " / " & r2.Rows.Count & "). 같은 줄 수로 다시 고르세요."
            Exit Function
        End If
    Else
        api.CoGuard_Notice "비교할 두 열을 골라야 합니다." & vbCrLf & "나란한 두 열을 한 덩어리로 드래그하거나, Ctrl을 누른 채 두 열을 따로 고르세요."
        Exit Function
    End If
    ' 열 전체를 골랐으면 사용 중인 줄까지만
    lastRow = Application.WorksheetFunction.Max(api.CoRange_LastRow(ws, r1.Column), api.CoRange_LastRow(ws, r2.Column))
    If lastRow < r1.Row Then api.CoGuard_Notice "선택한 열에 값이 없습니다.": Exit Function
    If r1.Row + r1.Rows.Count - 1 > lastRow Then
        n = lastRow - r1.Row + 1
        Set r1 = r1.Resize(n)
        Set r2 = r2.Resize(n)
    End If
    Set a = r1: Set b = r2
    PickTwo = True
End Function

' 값을 비교용 글자로(오류는 #로 시작하는 글자 그대로, 빈 칸은 "").
Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Then
        CellText = "#err" & CStr(CLng(v))
    ElseIf IsEmpty(v) Then
        CellText = ""
    Else
        CellText = CStr(v)
    End If
End Function
