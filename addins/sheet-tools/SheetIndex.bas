Attribute VB_Name = "SheetIndex"
'==============================================================
' 기능: 통합 문서의 모든 시트를 링크로 나열한 "목차" 시트를 맨 앞에 만든다(시트 이름, 상태, 사용 범위, 클릭하면 이동).
' 사용법: 시트 탭 우클릭 > 칼퇴 시트 도구 > 시트 목차 만들기. 이미 "목차" 시트가 있으면 내용을 새로 쓸지 묻는다.
' 설정: indexName(목차 시트 이름, 기본 "목차"), indexIncludeHidden(숨긴 시트도 넣기, 기본 켬).
' 의존: SheetTools.bas, 공용 라이브러리(CoSetting_*, CoGuard_*)
'==============================================================
Option Explicit
Option Private Module

Public Sub SheetIndex_Make()
    Dim api As Object, wb As Workbook, nm As String, withHidden As Boolean
    Dim idx As Worksheet, sh As Object, r As Long, tok As Long, msg As String, state As String, ur As String
    Set api = ST_Api()
    If api Is Nothing Then Exit Sub
    Set wb = ActiveWorkbook
    If wb Is Nothing Then api.CoGuard_Notice "열린 통합 문서가 없습니다.": Exit Sub
    If wb.ProtectStructure Then api.CoGuard_Notice "통합 문서 구조가 보호되어 있어 시트를 만들 수 없습니다.": Exit Sub
    nm = api.CoText_SafeSheetName(api.CoSetting_Text(ST_ADDIN_ID, "indexName", "목차"))
    withHidden = api.CoSetting_Bool(ST_ADDIN_ID, "indexIncludeHidden", True)
    If ST_SheetExists(wb, nm) Then
        If Not api.CoGuard_Ask("'" & nm & "' 시트가 이미 있습니다. 내용을 지우고 새로 쓸까요?" & vbCrLf & "(실행 후 Ctrl+Z로 되돌릴 수 없습니다.)") Then Exit Sub
        Set idx = wb.Worksheets(nm)
        If idx.ProtectContents Then api.CoGuard_Notice "'" & nm & "' 시트가 보호되어 있습니다.": Exit Sub
        idx.Cells.Clear
    Else
        tok = api.CoGuard_Begin(True, True)
        Set idx = wb.Worksheets.Add(Before:=wb.Sheets(1))
        idx.Name = nm
        api.CoGuard_End tok
    End If
    tok = api.CoGuard_Begin(True, True)
    On Error GoTo EH
    idx.Range("A1:D1").Value2 = Array("번호", "시트", "상태", "사용 범위")
    idx.Range("A1:D1").Font.Bold = True
    r = 1
    For Each sh In wb.Sheets
        If sh.Name <> nm Then
            If sh.Visible = xlSheetVisible Or withHidden Then
                r = r + 1
                idx.Cells(r, 1).Value2 = r - 1
                state = IIf(sh.Visible = xlSheetVisible, "", IIf(sh.Visible = xlSheetHidden, "숨김", "완전히 숨김"))
                idx.Cells(r, 3).Value2 = state
                ur = ""
                If TypeName(sh) = "Worksheet" Then
                    ur = sh.UsedRange.Address(False, False)
                Else
                    ur = "(차트 시트)"
                End If
                idx.Cells(r, 4).Value2 = ur
                If sh.Visible = xlSheetVisible Then
                    idx.Hyperlinks.Add Anchor:=idx.Cells(r, 2), Address:="", SubAddress:=ST_Quote(sh.Name) & "!A1", TextToDisplay:=sh.Name
                Else
                    idx.Cells(r, 2).Value2 = sh.Name        ' 숨긴 시트는 이동할 수 없어 링크 없이 이름만
                End If
            End If
        End If
    Next sh
    idx.Columns("A:D").AutoFit
    api.CoGuard_End tok
    idx.Activate
    idx.Range("A1").Select
    api.CoGuard_Notice "'" & nm & "' 시트에 " & (r - 1) & "개 시트를 정리했습니다."
    Exit Sub
EH:
    msg = "목차를 만드는 중 문제가 생겼습니다: " & Err.Description
    Err.Clear
    api.CoGuard_End tok
    api.CoGuard_Notice msg
End Sub
