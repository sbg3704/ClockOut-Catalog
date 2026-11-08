Attribute VB_Name = "FolderList"
'==============================================================
' 기능: 폴더를 골라 그 안의 파일 목록(이름, 확장자, 크기, 수정일, 위치)을 새 시트에 만든다. 이름을 누르면 파일이 열린다.
' 사용법: 시트 탭 우클릭 > 칼퇴 시트 도구 > 폴더 파일 목록 만들기.
' 설정: listSubfolders(하위 폴더까지, 기본 켬), listMaxRows(최대 줄 수, 기본 20000, 최대 100000).
' 규칙: 새 시트에 만든다(기존 시트는 건드리지 않음). 읽을 수 없는 폴더는 건너뛴다.
' 의존: SheetTools.bas, 공용 라이브러리(CoFile_PickFolder, CoSetting_*, CoGuard_*)
'==============================================================
Option Explicit
Option Private Module

Private mRows() As Variant
Private mCount As Long
Private mMax As Long
Private mCut As Boolean

Public Sub FolderList_Make()
    Dim api As Object, folder As String, deep As Boolean, wb As Workbook, ws As Worksheet, tok As Long, msg As String, i As Long
    Set api = ST_Api()
    If api Is Nothing Then Exit Sub
    Set wb = ActiveWorkbook
    If wb Is Nothing Then api.CoGuard_Notice "열린 통합 문서가 없습니다.": Exit Sub
    If wb.ProtectStructure Then api.CoGuard_Notice "통합 문서 구조가 보호되어 있어 시트를 만들 수 없습니다.": Exit Sub
    folder = api.CoFile_PickFolder("파일 목록을 만들 폴더", api.CoFile_WorkbookFolder(wb))
    If Len(folder) = 0 Then Exit Sub
    deep = api.CoSetting_Bool(ST_ADDIN_ID, "listSubfolders", True)
    mMax = CLng(api.CoSetting_Number(ST_ADDIN_ID, "listMaxRows", 20000, 100, 100000))
    mCount = 0: mCut = False
    ReDim mRows(1 To mMax, 1 To 6)
    tok = api.CoGuard_Begin(True, True)
    On Error GoTo EH
    Scan CreateObject("Scripting.FileSystemObject").GetFolder(folder), deep, folder
    Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = ST_UniqueSheetName(api, wb, "파일목록")
    ws.Range("A1:F1").Value2 = Array("이름", "확장자", "크기(KB)", "수정일", "하위 폴더", "전체 경로")
    ws.Range("A1:F1").Font.Bold = True
    If mCount > 0 Then
        ws.Range("A2").Resize(mCount, 6).Value2 = SliceRows()
        ws.Range("D2").Resize(mCount, 1).NumberFormat = "yyyy-mm-dd hh:mm"
        ws.Range("C2").Resize(mCount, 1).NumberFormat = "#,##0.0"
        For i = 1 To mCount
            ws.Hyperlinks.Add Anchor:=ws.Cells(i + 1, 1), Address:=CStr(mRows(i, 6)), TextToDisplay:=CStr(mRows(i, 1))
        Next i
    End If
    ws.Columns("A:F").AutoFit
    If ws.Columns("F").ColumnWidth > 70 Then ws.Columns("F").ColumnWidth = 70
    api.CoGuard_End tok
    ws.Activate
    api.CoGuard_Notice mCount & "개 파일을 '" & ws.Name & "' 시트에 정리했습니다." & IIf(mCut, vbCrLf & "(최대 " & mMax & "줄에서 멈췄습니다. 설정에서 늘릴 수 있습니다.)", "")
    Exit Sub
EH:
    msg = "목록을 만드는 중 문제가 생겼습니다: " & Err.Description
    Err.Clear
    api.CoGuard_End tok
    api.CoGuard_Notice msg
End Sub

Private Sub Scan(ByVal fld As Object, ByVal deep As Boolean, ByVal root As String)
    Dim f As Object, sub1 As Object, rel As String
    On Error Resume Next
    For Each f In fld.Files
        If mCount >= mMax Then mCut = True: Exit Sub
        mCount = mCount + 1
        mRows(mCount, 1) = f.Name
        mRows(mCount, 2) = LCase$(CreateObject("Scripting.FileSystemObject").GetExtensionName(f.Name))
        mRows(mCount, 3) = Round(f.Size / 1024, 1)
        mRows(mCount, 4) = f.DateLastModified
        rel = Mid$(fld.Path, Len(root) + 2)
        mRows(mCount, 5) = rel
        mRows(mCount, 6) = f.Path
    Next f
    If deep Then
        For Each sub1 In fld.SubFolders
            If mCut Then Exit Sub
            Scan sub1, deep, root
        Next sub1
    End If
End Sub

' 모은 줄만 잘라 2차원 배열로(1..mCount).
Private Function SliceRows() As Variant
    Dim out() As Variant, i As Long, j As Long
    ReDim out(1 To mCount, 1 To 6)
    For i = 1 To mCount
        For j = 1 To 6
            out(i, j) = mRows(i, j)
        Next j
    Next i
    SliceRows = out
End Function
