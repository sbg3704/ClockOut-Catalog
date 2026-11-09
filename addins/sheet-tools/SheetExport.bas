Attribute VB_Name = "SheetExport"
'==============================================================
' 기능: 시트마다 따로 .xlsx 파일을 만든다(원본 통합 문서는 바뀌지 않는다).
' 사용법: (선택) Ctrl로 내보낼 시트 탭만 고른 뒤, 시트 탭 우클릭 > 칼퇴 시트 도구 > 시트별 파일로 내보내기. 저장할 폴더를 고른다.
'         시트를 여러 개 고르지 않았으면 보이는 워크시트 전부를 내보낸다.
' 설정: exportValues(수식을 값으로 바꿔 저장, 기본 켬 - 다른 시트·외부 참조가 깨지지 않게), exportPrefix(파일 이름 앞에 원본 통합 문서 이름 붙이기).
' 규칙: 같은 이름의 파일이 있으면 덮어쓰지 않고 (2), (3)을 붙인다. 숨긴 시트와 차트 시트는 건너뛴다.
' 의존: SheetTools.bas, 공용 라이브러리(CoFile_PickFolder, CoText_SafeFileName, CoFile_Exists, CoGuard_*)
'==============================================================
Option Explicit
Option Private Module

Public Sub SheetExport_Each()
    Dim api As Object, wb As Workbook, list As Collection, sh As Object, folder As String
    Dim toValues As Boolean, usePrefix As Boolean, done As Long, failed As String, tok As Long
    Dim fileBase As String, path As String, n As Long, nb As Workbook, msg As String, i As Long
    Set api = ST_Api()
    If api Is Nothing Then Exit Sub
    Set wb = ActiveWorkbook
    If wb Is Nothing Then api.CoGuard_Notice "열린 통합 문서가 없습니다.": Exit Sub
    Set list = New Collection
    On Error Resume Next
    If ActiveWindow.SelectedSheets.Count > 1 Then
        For Each sh In ActiveWindow.SelectedSheets
            If TypeName(sh) = "Worksheet" Then list.Add sh
        Next sh
    End If
    On Error GoTo 0
    If list.Count = 0 Then
        For Each sh In wb.Worksheets
            If sh.Visible = xlSheetVisible Then list.Add sh
        Next sh
    End If
    If list.Count = 0 Then api.CoGuard_Notice "내보낼 시트가 없습니다.": Exit Sub
    toValues = api.CoSetting_Bool(ST_ADDIN_ID, "exportValues", True)
    usePrefix = api.CoSetting_Bool(ST_ADDIN_ID, "exportPrefix", False)
    folder = api.CoFile_PickFolder("시트별 파일을 저장할 폴더", api.CoFile_WorkbookFolder(wb))
    If Len(folder) = 0 Then Exit Sub
    If Not api.CoGuard_Ask(list.Count & "개 시트를 아래 폴더에 파일로 저장합니다." & vbCrLf & folder & vbCrLf & _
                           IIf(toValues, "(수식은 값으로 바꿔 저장)", "(수식 그대로 저장)") & vbCrLf & "진행할까요?") Then Exit Sub
    tok = api.CoGuard_Begin(True, True, False, True)
    On Error GoTo EH
    For i = 1 To list.Count
        Set sh = list(i)
        fileBase = api.CoText_SafeFileName(sh.Name)
        If usePrefix Then fileBase = api.CoText_SafeFileName(Replace(wb.Name, ".xls", "", 1, 1, vbTextCompare)) & "_" & fileBase
        path = api.CoText_PathJoin(folder, fileBase & ".xlsx")
        n = 1
        Do While api.CoFile_Exists(path)
            n = n + 1
            path = api.CoText_PathJoin(folder, fileBase & " (" & n & ").xlsx")
        Loop
        sh.Copy
        Set nb = ActiveWorkbook
        If toValues Then ToValuesAndUnlink nb
        nb.SaveAs Filename:=path, FileFormat:=51
        nb.Close SaveChanges:=False
        Set nb = Nothing
        done = done + 1
    Next i
    api.CoGuard_End tok
    wb.Activate
    api.CoGuard_Notice done & "개 시트를 파일로 저장했습니다." & vbCrLf & folder
    Exit Sub
EH:
    msg = "내보내는 중 문제가 생겼습니다(" & done & "개 저장 뒤): " & Err.Description
    Err.Clear
    On Error Resume Next
    If Not nb Is Nothing Then nb.Close SaveChanges:=False
    api.CoGuard_End tok
    wb.Activate
    api.CoGuard_Notice msg
End Sub

' 새 통합 문서의 수식을 값으로 바꾸고 외부 연결을 끊는다.
Private Sub ToValuesAndUnlink(ByVal nb As Workbook)
    Dim src As Variant, i As Long, c As Range
    On Error Resume Next
    With nb.Worksheets(1).UsedRange
        .Value2 = .Value2
        If Err.Number <> 0 Then                      ' 병합 칸 등으로 한 번에 안 되면 붙여넣기로
            Err.Clear
            .Copy
            .PasteSpecial xlPasteValues
            Application.CutCopyMode = False
        End If
    End With
    Err.Clear
    src = nb.LinkSources(xlExcelLinks)
    If IsArray(src) Then
        For i = LBound(src) To UBound(src)
            nb.BreakLink src(i), xlLinkTypeExcelLinks
        Next i
    End If
End Sub
