Attribute VB_Name = "BookCleanPage"
'==============================================================
' 기능: 지금 시트의 인쇄 설정(용지·방향·여백·배율·머리글/바닥글·가운데 맞춤)을 다른 시트들에 똑같이 적용한다.
' 사용법: 기준 시트를 열어 둔다. (선택) Ctrl로 적용할 시트 탭을 함께 고른다. 시트 탭 우클릭 > 칼퇴 문서 정리 > 인쇄 설정 일괄 복사.
'         다른 시트를 고르지 않았으면 이 통합 문서의 모든 워크시트에 적용한다.
' 설정: pageCopyTitles(반복할 제목 행·열도 복사, 기본 끔), pageCopyArea(인쇄 영역도 복사, 기본 끔 - 시트마다 달라 보통 끔).
' 의존: BookClean.bas, 공용 라이브러리(CoGuard_*, CoSetting_Bool)
'==============================================================
Option Explicit
Option Private Module

Public Sub BookClean_CopyPage()
    Dim api As Object, wb As Workbook, src As Worksheet, sh As Object, targets As Collection
    Dim withTitles As Boolean, withArea As Boolean, tok As Long, done As Long, msg As String, bad As String
    Set api = BC_Api()
    If api Is Nothing Then Exit Sub
    Set wb = ActiveWorkbook
    If wb Is Nothing Or TypeName(ActiveSheet) <> "Worksheet" Then api.CoGuard_Notice "기준이 될 워크시트를 먼저 여세요.": Exit Sub
    Set src = ActiveSheet
    withTitles = api.CoSetting_Bool(BC_ADDIN_ID, "pageCopyTitles", False)
    withArea = api.CoSetting_Bool(BC_ADDIN_ID, "pageCopyArea", False)
    Set targets = New Collection
    On Error Resume Next
    If ActiveWindow.SelectedSheets.Count > 1 Then
        For Each sh In ActiveWindow.SelectedSheets
            If TypeName(sh) = "Worksheet" Then If Not sh Is src Then targets.Add sh
        Next sh
    End If
    On Error GoTo 0
    If targets.Count = 0 Then
        For Each sh In wb.Worksheets
            If Not sh Is src Then If sh.Visible = xlSheetVisible Then targets.Add sh
        Next sh
    End If
    If targets.Count = 0 Then api.CoGuard_Notice "적용할 다른 시트가 없습니다.": Exit Sub
    If Not BC_Confirm(api, "'" & src.Name & "'의 인쇄 설정을 " & targets.Count & "개 시트에 복사합니다." & vbCrLf & _
                      "(용지·방향·여백·배율·머리글/바닥글·가운데 맞춤" & IIf(withTitles, "·반복 제목", "") & IIf(withArea, "·인쇄 영역", "") & ")") Then Exit Sub
    tok = api.CoGuard_Begin(True, True)
    On Error GoTo EH
    For Each sh In targets
        If sh.ProtectContents Then
            bad = bad & vbCrLf & "  - " & sh.Name & " (보호됨)"
        Else
            CopyOne src, sh, withTitles, withArea
            done = done + 1
        End If
    Next sh
    api.CoGuard_End tok
    api.CoGuard_Notice done & "개 시트에 인쇄 설정을 복사했습니다." & IIf(Len(bad) > 0, vbCrLf & "건너뜀:" & bad, "")
    Exit Sub
EH:
    msg = "복사하는 중 문제가 생겼습니다: " & Err.Description
    Err.Clear
    api.CoGuard_End tok
    api.CoGuard_Notice msg
End Sub

' 한 시트에서 다른 시트로. 프린터 드라이버가 못 받는 항목은 건너뛴다(항목별로 오류 무시).
Private Sub CopyOne(ByVal a As Worksheet, ByVal b As Worksheet, ByVal withTitles As Boolean, ByVal withArea As Boolean)
    On Error Resume Next
    With b.PageSetup
        .PaperSize = a.PageSetup.PaperSize
        .Orientation = a.PageSetup.Orientation
        .LeftMargin = a.PageSetup.LeftMargin
        .RightMargin = a.PageSetup.RightMargin
        .TopMargin = a.PageSetup.TopMargin
        .BottomMargin = a.PageSetup.BottomMargin
        .HeaderMargin = a.PageSetup.HeaderMargin
        .FooterMargin = a.PageSetup.FooterMargin
        .CenterHorizontally = a.PageSetup.CenterHorizontally
        .CenterVertically = a.PageSetup.CenterVertically
        .LeftHeader = a.PageSetup.LeftHeader
        .CenterHeader = a.PageSetup.CenterHeader
        .RightHeader = a.PageSetup.RightHeader
        .LeftFooter = a.PageSetup.LeftFooter
        .CenterFooter = a.PageSetup.CenterFooter
        .RightFooter = a.PageSetup.RightFooter
        .PrintGridlines = a.PageSetup.PrintGridlines
        .PrintHeadings = a.PageSetup.PrintHeadings
        .Order = a.PageSetup.Order
        If a.PageSetup.Zoom = False Then
            .Zoom = False
            .FitToPagesWide = a.PageSetup.FitToPagesWide
            .FitToPagesTall = a.PageSetup.FitToPagesTall
        Else
            .Zoom = a.PageSetup.Zoom
        End If
        If withTitles Then
            .PrintTitleRows = a.PageSetup.PrintTitleRows
            .PrintTitleColumns = a.PageSetup.PrintTitleColumns
        End If
        If withArea Then .PrintArea = a.PageSetup.PrintArea
    End With
End Sub
