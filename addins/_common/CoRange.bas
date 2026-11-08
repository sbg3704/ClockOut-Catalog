Attribute VB_Name = "CoRange"
'==============================================================
' 기능: 칼퇴 공용 라이브러리 - 시트·선택 범위 안전 검사와 범위 도우미.
'       검사 함수는 안내 창을 직접 띄우지 않고 reason(이유 글)을 돌려준다. 띄울지는 매크로가 정한다.
' 사용법:
'   Dim ws As Worksheet, rng As Range, why As String
'   If Not CoRange_GetSelection(ws, rng, why) Then CoGuard_Notice why: Exit Sub
'   If Not CoRange_CanEdit(rng) Then CoGuard_Notice "보호된 셀입니다.": Exit Sub
'   last = CoRange_LastRow(ws, 1)            ' A열 마지막 값 있는 행(없으면 0)
' 의존: 없음(Excel 개체만). image-fit의 GetTargetRange·IsSheetBlocked 검사를 일반화했다.
'==============================================================
Option Explicit
Option Private Module

' 활성 시트가 워크시트인지 확인한다.
Public Function CoRange_ActiveWorksheet(ByRef ws As Worksheet, ByRef reason As String) As Boolean
    reason = ""
    If ActiveWorkbook Is Nothing Then reason = "열려 있는 통합 문서가 없습니다.": Exit Function
    If TypeName(ActiveSheet) <> "Worksheet" Then
        reason = "워크시트에서만 쓸 수 있습니다. (차트 시트 등은 안 됩니다)"
        Exit Function
    End If
    Set ws = ActiveSheet
    CoRange_ActiveWorksheet = True
End Function

' 선택한 셀 범위를 안전하게 얻는다.
'   allowMulti=False면 떨어진 여러 영역 중 첫 덩어리만 돌려준다.
'   allowWhole=False면 행·열 전체 선택을 막는다(수백만 칸을 도는 사고 방지).
Public Function CoRange_GetSelection(ByRef ws As Worksheet, ByRef rng As Range, ByRef reason As String, _
                                     Optional ByVal allowMulti As Boolean = False, _
                                     Optional ByVal allowWhole As Boolean = False) As Boolean
    If Not CoRange_ActiveWorksheet(ws, reason) Then Exit Function
    If TypeName(Selection) <> "Range" Then reason = "셀을 먼저 선택하세요.": Exit Function
    If Selection.Areas.Count > 1 And Not allowMulti Then
        Set rng = Selection.Areas(1)
    Else
        Set rng = Selection
    End If
    If Not allowWhole Then
        If CoRange_IsWholeRowOrCol(rng) Then
            reason = "행 또는 열 전체가 선택되어 있습니다. 필요한 칸만 선택하세요."
            Exit Function
        End If
    End If
    CoRange_GetSelection = True
End Function

' 어느 한 덩어리라도 행 전체나 열 전체면 True.
Public Function CoRange_IsWholeRowOrCol(ByVal rng As Range) As Boolean
    Dim a As Range
    For Each a In rng.Areas
        If a.Rows.Count >= rng.Worksheet.Rows.Count Or a.Columns.Count >= rng.Worksheet.Columns.Count Then
            CoRange_IsWholeRowOrCol = True
            Exit Function
        End If
    Next a
End Function

' 숨긴 행·열이 아니고 크기가 있으면 True(사진·도형을 놓을 수 있는 칸).
Public Function CoRange_IsVisible(ByVal rng As Range) As Boolean
    CoRange_IsVisible = (rng.Width > 0 And rng.Height > 0)
End Function

' 첫 칸이 병합 셀이면 병합 영역 전체, 아니면 그 칸.
Public Function CoRange_FirstCellArea(ByVal rng As Range) As Range
    Set CoRange_FirstCellArea = rng.Cells(1, 1).MergeArea
End Function

' 범위에 병합 셀이 하나라도 있으면 True.
Public Function CoRange_HasMerged(ByVal rng As Range) As Boolean
    Dim m As Variant
    m = rng.MergeCells                                ' 섞여 있으면 Null
    If IsNull(m) Then CoRange_HasMerged = True Else CoRange_HasMerged = CBool(m)
End Function

' 시트 보호 여부. forShapes=True면 개체(사진·도형) 보호를 본다.
Public Function CoRange_IsProtected(ByVal ws As Worksheet, Optional ByVal forShapes As Boolean = False) As Boolean
    If forShapes Then
        CoRange_IsProtected = ws.ProtectDrawingObjects
    Else
        CoRange_IsProtected = ws.ProtectContents
    End If
End Function

' 값을 써도 되는 범위면 True(시트 보호 중이면 잠긴 칸이 하나도 없어야 한다).
Public Function CoRange_CanEdit(ByVal rng As Range) As Boolean
    Dim lk As Variant
    If Not rng.Worksheet.ProtectContents Then CoRange_CanEdit = True: Exit Function
    lk = rng.Locked                                   ' 섞여 있으면 Null
    If IsNull(lk) Then Exit Function
    CoRange_CanEdit = Not CBool(lk)
End Function

' 값(수식 포함)이 있는 마지막 행. col을 주면 그 열만 본다. 비어 있으면 0.
' 주의: col 없이 부르면 Find를 쓰므로 사용자의 찾기 창 옵션(LookIn 등)이 바뀔 수 있다.
Public Function CoRange_LastRow(ByVal ws As Worksheet, Optional ByVal col As Long = 0) As Long
    Dim c As Range
    If col > 0 Then
        Set c = ws.Cells(ws.Rows.Count, col)
        If Len(CStr(c.Formula)) = 0 Then Set c = c.End(xlUp)   ' 맨 아래 칸에 값이 있으면 그대로
        If Len(CStr(c.Formula)) > 0 Then CoRange_LastRow = c.Row
    Else
        Set c = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                              SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        If Not c Is Nothing Then CoRange_LastRow = c.Row
    End If
End Function

' 값(수식 포함)이 있는 마지막 열. row를 주면 그 행만 본다. 비어 있으면 0.
Public Function CoRange_LastCol(ByVal ws As Worksheet, Optional ByVal row As Long = 0) As Long
    Dim c As Range
    If row > 0 Then
        Set c = ws.Cells(row, ws.Columns.Count)
        If Len(CStr(c.Formula)) = 0 Then Set c = c.End(xlToLeft)
        If Len(CStr(c.Formula)) > 0 Then CoRange_LastCol = c.Column
    Else
        Set c = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                              SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
        If Not c Is Nothing Then CoRange_LastCol = c.Column
    End If
End Function

' 범위 값을 항상 2차원 배열(1부터)로 돌려준다. 한 칸이어도 (1,1) 배열. 여러 덩어리면 첫 덩어리.
' 칸을 하나씩 읽고 쓰는 것보다 훨씬 빠르다. 다 고친 뒤 rng.Value = 배열로 한 번에 쓴다.
Public Function CoRange_Values2D(ByVal rng As Range) As Variant
    Dim a As Range, v As Variant
    Set a = rng.Areas(1)
    If a.Rows.Count = 1 And a.Columns.Count = 1 Then
        ReDim v(1 To 1, 1 To 1)
        v(1, 1) = a.Value
        CoRange_Values2D = v
    Else
        CoRange_Values2D = a.Value
    End If
End Function
