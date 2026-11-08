Attribute VB_Name = "BookFiles"
'==============================================================
' 기능: 파일·양식 정리 4개. 공유 전 정리 / 원클릭 백업 / 선택한 시트를 PDF 하나로 / 입력 칸 비우기(수식 유지).
' 사용법: 시트 탭 우클릭 > 칼퇴 통합 문서 도구 > 도구. (선택) Ctrl로 시트 탭 여러 장을 고르면 그 시트들만 대상.
' 설정: prepZoom(공유 전 정리 배율, 기본 100), clearOnlyUnlocked(입력 칸 비우기를 잠금 해제된 칸만, 기본 끔).
' 규칙: 백업은 저장된 파일만 되고 원본 옆에 `이름_20261108_1530.확장자`로 만든다(덮어쓰지 않음). 열려 있는 통합 문서는 바뀌지 않는다.
'       PDF는 원본 옆에 `이름_시트묶음_날짜시각.pdf`로 만든다. 입력 칸 비우기는 상수(숫자·글자)만 지우고 수식·서식은 둔다.
' 의존: BookTools.bas, 공용 라이브러리(CoFile_Exists, CoText_SafeFileName, CoSetting_*, CoGuard_*)
'==============================================================
Option Explicit
Option Private Module

'--- 공유 전 정리 -----------------------------------------------
Public Sub BookTools_PrepShare()
    Dim api As Object, wb As Workbook, sh As Object, z As Long, first As Object, n As Long, cur As Object
    Set api = BT_Api()
    If api Is Nothing Then Exit Sub
    Set wb = ActiveWorkbook
    If wb Is Nothing Then api.CoGuard_Notice "열린 통합 문서가 없습니다.": Exit Sub
    z = CLng(api.CoSetting_Number(BT_ADDIN_ID, "prepZoom", 100, 10, 400))
    If Not api.CoGuard_Ask("보이는 모든 시트의 선택을 A1로, 스크롤을 맨 위로, 배율을 " & z & "%로 맞추고 첫 시트를 엽니다." & vbCrLf & "진행할까요?") Then Exit Sub
    Set cur = ActiveSheet
    Application.ScreenUpdating = False
    On Error Resume Next
    For Each sh In wb.Sheets
        If sh.Visible = xlSheetVisible Then
            sh.Activate
            If TypeName(sh) = "Worksheet" Then
                ActiveWindow.ScrollRow = 1
                ActiveWindow.ScrollColumn = 1
                sh.Range("A1").Select
                ActiveWindow.Zoom = z
            End If
            If first Is Nothing Then Set first = sh
            n = n + 1
        End If
    Next sh
    If Not first Is Nothing Then first.Activate
    Application.ScreenUpdating = True
    On Error GoTo 0
    api.CoGuard_Notice n & "개 시트를 정리했습니다. 저장하면 이 모양으로 열립니다."
End Sub

'--- 원클릭 백업 ------------------------------------------------
Public Sub BookTools_Backup()
    Dim api As Object, wb As Workbook, base As String, ext As String, p As String, dotPos As Long, path As String, n As Long
    Set api = BT_Api()
    If api Is Nothing Then Exit Sub
    Set wb = ActiveWorkbook
    If wb Is Nothing Then api.CoGuard_Notice "열린 통합 문서가 없습니다.": Exit Sub
    If Len(wb.Path) = 0 Then api.CoGuard_Notice "아직 저장하지 않은 문서입니다. 먼저 저장하세요.": Exit Sub
    dotPos = InStrRev(wb.Name, ".")
    If dotPos > 0 Then base = Left$(wb.Name, dotPos - 1): ext = Mid$(wb.Name, dotPos) Else base = wb.Name
    p = wb.Path & Application.PathSeparator & base & "_" & Format$(Now, "yyyymmdd_hhnn")
    path = p & ext
    Do While api.CoFile_Exists(path)
        n = n + 1
        path = p & "_" & (n + 1) & ext
    Loop
    On Error GoTo EH
    wb.SaveCopyAs path
    api.CoGuard_Notice "백업했습니다(현재 문서의 마지막 저장 상태가 아니라 지금 화면 상태):" & vbCrLf & path
    Exit Sub
EH:
    api.CoGuard_Notice "백업하지 못했습니다: " & Err.Description
End Sub

'--- 선택한 시트를 PDF 하나로 ------------------------------------
Public Sub BookTools_ExportPdf()
    Dim api As Object, wb As Workbook, sel As Collection, arr() As String, i As Long, sh As Object, base As String, path As String, n As Long, cur As Object
    Set api = BT_Api()
    If api Is Nothing Then Exit Sub
    Set wb = ActiveWorkbook
    If wb Is Nothing Then api.CoGuard_Notice "열린 통합 문서가 없습니다.": Exit Sub
    If Len(wb.Path) = 0 Then api.CoGuard_Notice "아직 저장하지 않은 문서입니다. 먼저 저장하세요.": Exit Sub
    Set sel = BT_SelectedSheets()
    If sel.Count = 1 Then
        Set sel = New Collection
        For Each sh In wb.Worksheets
            If sh.Visible = xlSheetVisible Then sel.Add sh
        Next sh
    End If
    If sel.Count = 0 Then api.CoGuard_Notice "내보낼 시트가 없습니다.": Exit Sub
    base = api.CoText_SafeFileName(Left$(wb.Name, IIf(InStrRev(wb.Name, ".") > 0, InStrRev(wb.Name, ".") - 1, Len(wb.Name))))
    path = wb.Path & Application.PathSeparator & base & "_시트묶음_" & Format$(Now, "yyyymmdd_hhnn") & ".pdf"
    Do While api.CoFile_Exists(path)
        n = n + 1
        path = wb.Path & Application.PathSeparator & base & "_시트묶음_" & Format$(Now, "yyyymmdd_hhnn") & "_" & (n + 1) & ".pdf"
    Loop
    If Not api.CoGuard_Ask(sel.Count & "개 시트를 PDF 파일 하나로 저장합니다." & vbCrLf & path & vbCrLf & "진행할까요?") Then Exit Sub
    ReDim arr(1 To sel.Count)
    For i = 1 To sel.Count: arr(i) = sel(i).Name: Next i
    Set cur = ActiveSheet
    On Error GoTo EH
    wb.Sheets(arr).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path, Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
    cur.Select
    api.CoGuard_Notice "PDF로 저장했습니다:" & vbCrLf & path
    Exit Sub
EH:
    Dim msg As String
    msg = "PDF로 저장하지 못했습니다: " & Err.Description
    Err.Clear
    On Error Resume Next
    cur.Select
    api.CoGuard_Notice msg
End Sub

'--- 입력 칸 비우기(수식 유지) ------------------------------------
Public Sub BookTools_ClearInputs()
    Dim api As Object, wb As Workbook, sel As Collection, sh As Object, onlyUnlocked As Boolean, rng As Range, cel As Range
    Dim n As Long, list As New Collection, tok As Long, msg As String, skipped As String, cnt As Long, o As Object
    Set api = BT_Api()
    If api Is Nothing Then Exit Sub
    Set wb = ActiveWorkbook
    If wb Is Nothing Then api.CoGuard_Notice "열린 통합 문서가 없습니다.": Exit Sub
    onlyUnlocked = api.CoSetting_Bool(BT_ADDIN_ID, "clearOnlyUnlocked", False)
    Set sel = BT_SelectedSheets()
    For Each sh In sel
        If TypeName(sh) = "Worksheet" Then
            If sh.ProtectContents Then
                skipped = skipped & vbCrLf & "  - " & sh.Name & " (보호됨)"
            Else
                Set rng = Nothing
                On Error Resume Next
                Set rng = sh.UsedRange.SpecialCells(xlCellTypeConstants)
                On Error GoTo 0
                If Not rng Is Nothing Then
                    If onlyUnlocked Then
                        Dim keep As Range
                        Set keep = Nothing
                        For Each cel In rng.Cells
                            If Not cel.Locked Then
                                If keep Is Nothing Then Set keep = cel Else Set keep = Union(keep, cel)
                            End If
                        Next cel
                        Set rng = keep
                    End If
                    If Not rng Is Nothing Then list.Add rng: n = n + rng.CountLarge
                End If
            End If
        End If
    Next sh
    If n = 0 Then api.CoGuard_Notice "비울 입력 칸(상수)이 없습니다." & skipped: Exit Sub
    If Not BT_Confirm(api, "시트 " & list.Count & "개에서 상수(숫자·글자) " & n & "칸을 비웁니다. 수식과 서식은 그대로 둡니다." & IIf(onlyUnlocked, vbCrLf & "(잠금 해제된 칸만)", "") & IIf(Len(skipped) > 0, vbCrLf & "건너뜀:" & skipped, "")) Then Exit Sub
    tok = api.CoGuard_Begin(True, True)
    On Error GoTo EH
    For Each o In list
        o.ClearContents
    Next o
    api.CoGuard_End tok
    api.CoGuard_Notice n & "칸을 비웠습니다."
    Exit Sub
EH:
    msg = "비우는 중 문제가 생겼습니다: " & Err.Description
    Err.Clear
    api.CoGuard_End tok
    api.CoGuard_Notice msg
End Sub
