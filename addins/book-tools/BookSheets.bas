Attribute VB_Name = "BookSheets"
'==============================================================
' 기능: 시트 탭 정리 4개. 숨은 시트 모두 보이기 / 선택한 시트만 남기고 숨기기 / 탭 이름순(자연 순서) 정렬 / 시트 이름 일괄 바꾸기.
' 사용법: 시트 탭 우클릭 > 칼퇴 통합 문서 도구 > 도구. (선택) Ctrl로 시트 탭 여러 장을 함께 고르면 그 시트들에만 적용한다.
' 설정: showVeryHidden(숨은 시트 보이기에 '완전히 숨김' 시트도 포함, 기본 끔).
' 규칙: 구조가 보호된 통합 문서는 하지 않는다. 이름 바꾸기는 금지 문자(: \ / ? * [ ])를 _로, 31자로 자르고 겹치면 건너뛴다.
' 의존: BookTools.bas, 공용 라이브러리(CoText_NaturalCompare, CoText_SafeSheetName, CoGuard_*, CoSetting_Bool)
'==============================================================
Option Explicit
Option Private Module

Private Function CanChange(ByVal api As Object, ByVal wb As Workbook) As Boolean
    If wb Is Nothing Then api.CoGuard_Notice "열린 통합 문서가 없습니다.": Exit Function
    If wb.ProtectStructure Then api.CoGuard_Notice "통합 문서 구조가 보호되어 있어 시트를 바꿀 수 없습니다.": Exit Function
    CanChange = True
End Function

'--- 숨은 시트 모두 보이기 ----------------------------------------
Public Sub BookTools_ShowAll()
    Dim api As Object, wb As Workbook, sh As Object, incVery As Boolean, n As Long, names As String
    Set api = BT_Api()
    If api Is Nothing Then Exit Sub
    Set wb = ActiveWorkbook
    If Not CanChange(api, wb) Then Exit Sub
    incVery = api.CoSetting_Bool(BT_ADDIN_ID, "showVeryHidden", False)
    For Each sh In wb.Sheets
        If sh.Visible = xlSheetHidden Or (incVery And sh.Visible = xlSheetVeryHidden) Then
            n = n + 1
            If n <= 10 Then names = names & vbCrLf & "  - " & sh.Name
        End If
    Next sh
    If n = 0 Then api.CoGuard_Notice "숨은 시트가 없습니다." & IIf(incVery, "", vbCrLf & "('완전히 숨김' 시트는 설정을 켜야 보입니다.)"): Exit Sub
    If Not api.CoGuard_Ask(n & "개 숨은 시트를 보이게 합니다:" & names & IIf(n > 10, vbCrLf & "  ... 외 " & (n - 10) & "개", "") & vbCrLf & "진행할까요?") Then Exit Sub
    For Each sh In wb.Sheets
        If sh.Visible = xlSheetHidden Or (incVery And sh.Visible = xlSheetVeryHidden) Then sh.Visible = xlSheetVisible
    Next sh
    api.CoGuard_Notice n & "개 시트를 보이게 했습니다."
End Sub

'--- 선택한 시트만 남기고 숨기기 -----------------------------------
Public Sub BookTools_HideOthers()
    Dim api As Object, wb As Workbook, keep As Collection, sh As Object, k As Object, isKeep As Boolean, n As Long, vis As Long
    Set api = BT_Api()
    If api Is Nothing Then Exit Sub
    Set wb = ActiveWorkbook
    If Not CanChange(api, wb) Then Exit Sub
    Set keep = BT_SelectedSheets()
    For Each sh In wb.Sheets
        isKeep = False
        For Each k In keep
            If k.Name = sh.Name Then isKeep = True: Exit For
        Next k
        If Not isKeep And sh.Visible = xlSheetVisible Then n = n + 1
    Next sh
    If n = 0 Then api.CoGuard_Notice "숨길 다른 시트가 없습니다.": Exit Sub
    If Not api.CoGuard_Ask("선택한 시트 " & keep.Count & "개만 남기고 나머지 보이는 시트 " & n & "개를 숨깁니다." & vbCrLf & "(숨은 시트는 '숨은 시트 모두 보이기'로 되돌릴 수 있습니다.) 진행할까요?") Then Exit Sub
    For Each sh In wb.Sheets
        isKeep = False
        For Each k In keep
            If k.Name = sh.Name Then isKeep = True: Exit For
        Next k
        If Not isKeep And sh.Visible = xlSheetVisible Then sh.Visible = xlSheetHidden
    Next sh
    api.CoGuard_Notice n & "개 시트를 숨겼습니다."
End Sub

'--- 시트 탭 이름순 정렬 ---------------------------------------------
Public Sub BookTools_SortTabs()
    Dim api As Object, wb As Workbook, names() As String, i As Long, j As Long, n As Long, t As String, cur As Object, tok As Long
    Set api = BT_Api()
    If api Is Nothing Then Exit Sub
    Set wb = ActiveWorkbook
    If Not CanChange(api, wb) Then Exit Sub
    n = wb.Sheets.Count
    If n < 2 Then api.CoGuard_Notice "정렬할 시트가 없습니다.": Exit Sub
    ReDim names(1 To n)
    For i = 1 To n: names(i) = wb.Sheets(i).Name: Next i
    For i = 2 To n                                          ' 삽입 정렬(안정, 자연 순서: 1호 < 2호 < 10호)
        t = names(i): j = i - 1
        Do While j >= 1
            If api.CoText_NaturalCompare(names(j), t) <= 0 Then Exit Do
            names(j + 1) = names(j): j = j - 1
        Loop
        names(j + 1) = t
    Next i
    If Not api.CoGuard_Ask("시트 " & n & "개의 탭을 이름순(1호, 2호, 10호 순서)으로 다시 늘어놓습니다." & vbCrLf & "(숨은 시트도 이름순에 포함됩니다.) 진행할까요?") Then Exit Sub
    Set cur = ActiveSheet
    tok = api.CoGuard_Begin(True, True)
    On Error Resume Next
    For i = n To 1 Step -1                                   ' 맨 뒤부터 이름순으로 앞으로 보낸다
        wb.Sheets(names(i)).Move Before:=wb.Sheets(1)
    Next i
    On Error GoTo 0
    cur.Activate
    api.CoGuard_End tok
    api.CoGuard_Notice "시트 " & n & "개를 이름순으로 정렬했습니다."
End Sub

'--- 시트 이름 일괄 바꾸기 -------------------------------------------
Public Sub BookTools_RenameSheets()
    Dim api As Object, wb As Workbook, targets As Collection, sh As Object, f As String, r As String, nn As String
    Dim plan As New Collection, n As Long, skipped As Long, o As Variant, cand As String
    Set api = BT_Api()
    If api Is Nothing Then Exit Sub
    Set wb = ActiveWorkbook
    If Not CanChange(api, wb) Then Exit Sub
    Set targets = BT_SelectedSheets()
    If targets.Count = 1 Then
        Set targets = New Collection
        For Each sh In wb.Sheets
            If sh.Visible = xlSheetVisible Then targets.Add sh
        Next sh
    End If
    If Not api.CoGuard_AskText("바꿀 글자를 입력하세요(비우면 모든 시트 이름 '앞'에 붙이기 모드)." & vbCrLf & "대상: " & targets.Count & "개 시트", f) Then Exit Sub
    If Not api.CoGuard_AskText(IIf(Len(f) = 0, "이름 앞에 붙일 글자", "'" & f & "' 대신 넣을 글자(비우면 지움)"), r) Then Exit Sub
    If Len(f) = 0 And Len(r) = 0 Then api.CoGuard_Notice "바꿀 내용이 없습니다.": Exit Sub
    For Each sh In targets
        If Len(f) = 0 Then nn = r & sh.Name Else nn = Replace(sh.Name, f, r, 1, -1, vbTextCompare)
        nn = api.CoText_SafeSheetName(nn)
        If nn <> sh.Name Then
            plan.Add Array(sh.Name, nn)
        End If
    Next sh
    If plan.Count = 0 Then api.CoGuard_Notice "이름이 바뀌는 시트가 없습니다.": Exit Sub
    If Not api.CoGuard_Ask(plan.Count & "개 시트의 이름을 바꿉니다. 예: '" & plan(1)(0) & "' -> '" & plan(1)(1) & "'" & vbCrLf & "(겹치는 이름은 건너뜁니다.) 진행할까요?") Then Exit Sub
    On Error Resume Next
    For Each o In plan
        cand = o(1)
        If Not SheetNameTaken(wb, cand, o(0)) Then
            wb.Sheets(o(0)).Name = cand
            If Err.Number = 0 Then n = n + 1 Else skipped = skipped + 1: Err.Clear
        Else
            skipped = skipped + 1
        End If
    Next o
    On Error GoTo 0
    api.CoGuard_Notice n & "개 시트 이름을 바꿨습니다." & IIf(skipped > 0, vbCrLf & "(" & skipped & "개는 이름이 겹쳐 건너뜀)", "")
End Sub

Private Function SheetNameTaken(ByVal wb As Workbook, ByVal nm As String, ByVal exceptName As String) As Boolean
    Dim sh As Object
    For Each sh In wb.Sheets
        If StrComp(sh.Name, nm, vbTextCompare) = 0 And StrComp(sh.Name, exceptName, vbTextCompare) <> 0 Then SheetNameTaken = True: Exit Function
    Next sh
End Function
