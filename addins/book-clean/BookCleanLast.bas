Attribute VB_Name = "BookCleanLast"
'==============================================================
' 기능: Ctrl+End이 엉뚱하게 아래·오른쪽 끝으로 가거나 파일이 무거울 때, 값 없이 서식만 남은 빈 줄·빈 열을 지워 마지막 셀을 바로잡는다.
' 사용법: 시트 탭 우클릭 > 칼퇴 문서 정리 > 마지막 셀 바로잡기.
' 설정: lastAllSheets(모든 시트를 한 번에, 기본 끔 - 지금 시트만).
' 규칙: 값·수식이 있는 마지막 칸 뒤쪽의 줄·열만 지운다. 보호된 시트와 도형·차트·표가 그 구역에 걸친 시트는 건너뛴다.
'       (그 구역의 서식·빈 칸 모양은 사라진다.) 지우기 전에 시트별 범위를 보여 주고 확인받는다.
' 의존: BookClean.bas, 공용 라이브러리(CoRange_LastRow, CoRange_LastCol, CoSetting_Bool, CoGuard_*)
'==============================================================
Option Explicit
Option Private Module

Public Sub BookClean_FixLastCell()
    Dim api As Object, wb As Workbook, ws As Worksheet, list As Collection, rep As String, tok As Long, msg As String
    Dim allSheets As Boolean, lr As Long, lc As Long, ur As Range, endR As Long, endC As Long, n As Long, info As Variant
    Dim skipped As String, cnt As Long
    Set api = BC_Api()
    If api Is Nothing Then Exit Sub
    Set wb = ActiveWorkbook
    If wb Is Nothing Then api.CoGuard_Notice "열린 통합 문서가 없습니다.": Exit Sub
    allSheets = api.CoSetting_Bool(BC_ADDIN_ID, "lastAllSheets", False)
    Set list = New Collection
    For Each ws In wb.Worksheets
        If allSheets Or ws Is ActiveSheet Then
            If ws.Visible = xlSheetVisible Or allSheets Then
                If ws.ProtectContents Or ws.ProtectDrawingObjects Then
                    skipped = skipped & vbCrLf & "  - " & ws.Name & " (보호됨)"
                ElseIf ws.ListObjects.Count > 0 Or ws.Shapes.Count > 0 Then
                    skipped = skipped & vbCrLf & "  - " & ws.Name & " (표·도형·차트 있음)"
                Else
                    Set ur = ws.UsedRange
                    endR = ur.Row + ur.Rows.Count - 1
                    endC = ur.Column + ur.Columns.Count - 1
                    lr = api.CoRange_LastRow(ws)
                    lc = api.CoRange_LastCol(ws)
                    If lr = 0 Then lr = 1
                    If lc = 0 Then lc = 1
                    If endR > lr Or endC > lc Then
                        list.Add Array(ws, lr, lc, endR, endC)
                        rep = rep & vbCrLf & "  - " & ws.Name & ": 사용 범위 끝 " & ws.Cells(endR, endC).Address(False, False) & _
                              " -> 실제 데이터 끝 " & ws.Cells(lr, lc).Address(False, False)
                    End If
                End If
            End If
        End If
    Next ws
    If list.Count = 0 Then
        api.CoGuard_Notice "바로잡을 시트가 없습니다. 마지막 셀이 이미 실제 데이터 끝입니다." & skipped
        Exit Sub
    End If
    If Not BC_Confirm(api, "값이 없는 아래쪽 줄·오른쪽 열(서식만 남은 곳)을 지웁니다:" & rep & IIf(Len(skipped) > 0, vbCrLf & vbCrLf & "건너뜀:" & skipped, "")) Then Exit Sub
    tok = api.CoGuard_Begin(True, True)
    On Error GoTo EH
    For Each info In list
        Set ws = info(0)
        If info(3) > info(1) Then ws.Range(ws.Rows(info(1) + 1), ws.Rows(info(3))).Delete
        If info(4) > info(2) Then ws.Range(ws.Columns(info(2) + 1), ws.Columns(info(4))).Delete
        n = ws.UsedRange.Rows.Count                  ' 사용 범위를 다시 계산시킨다
        cnt = cnt + 1
    Next info
    api.CoGuard_End tok
    api.CoGuard_Notice cnt & "개 시트의 마지막 셀을 바로잡았습니다. 저장하면 파일이 가벼워집니다."
    Exit Sub
EH:
    msg = "정리하는 중 문제가 생겼습니다: " & Err.Description
    Err.Clear
    api.CoGuard_End tok
    api.CoGuard_Notice msg
End Sub
