Attribute VB_Name = "SheetNav"
'@clockout id=sheet-nav name=시트 사이드바 host=excel
'@desc 엑셀 창 왼쪽이나 오른쪽에 시트 목록을 띄워 두고, 이름을 클릭해 바로 이동합니다.
'==============================================================
' 기능: 엑셀 창 가장자리에 세로로 붙는 시트 목록 창(사이드바 모양)을 띄운다. 목록에서 시트 이름을 누르면 그 시트로
'       이동하고, 시트가 많으면 위의 검색 칸으로 이름을 걸러 찾는다. 시트를 옮기거나 추가하면 목록이 따라 바뀐다.
' 매크로:
'   SheetNav_Toggle   사이드바 열기/닫기 (기본 단축키 없음, 설정에서 지정)
' 사용법: 리본 단추나 설정에서 지정한 단축키로 연다. 창 위 "반대쪽" 단추로 왼쪽/오른쪽을 바꾼다.
'         쪽·너비·글자 크기·숨김 시트 표시·단축키는 칼퇴 설정창에서 바꾼다
'         (선언: sheet-nav.clockout.json, 값 저장: %APPDATA%\ClockOut\settings\sheet-nav.json).
' 규칙: 화면은 SheetNavForm.frm, 이벤트는 SheetNavEvents.cls, 목록 만들기·이동·자리 잡기는 이 파일에 있다.
'       아주 숨김(VeryHidden) 시트는 보이지 않는다. 숨김 시트는 설정을 켜면 "(숨김)"으로 보이고, 누르면 보이게 할지 묻는다.
' 한계: VBA로는 엑셀 안에 붙는 진짜 작업창을 만들 수 없어, 엑셀 창 가장자리에 맞춰 놓는 떠 있는 창이다.
'       엑셀 창 크기를 바꾸면 다시 맞추고, 창을 손으로 옮기면 그 자리에 둔다.
' 의존: 공용 라이브러리 clockout-lib 0.1.001 이상(CoSetting_*, CoGuard_Ask). CoLibClient.bas가 연결한다.
'==============================================================
Option Explicit

Private Const ADDIN_ID As String = "sheet-nav"      ' .clockout.json의 id와 같게
Private Const ADDIN_NAME As String = "시트 사이드바"
Private Const MIN_LIB As String = "0.1.001"          ' .clockout.json의 dependencies.min_version과 같게
Private Const DEFAULT_KEY As String = ""
Private Const SIDE_RIGHT As String = "오른쪽"
Private Const SIDE_LEFT As String = "왼쪽"
Private Const MARGIN_PT As Double = 12

Private mApi As Object
Private mForm As SheetNavForm
Private mEvents As SheetNavEvents
Private mOn As Boolean
Private mRight As Boolean
Private mWidth As Double
Private mFontPt As Double
Private mShowHidden As Boolean
Private mKey As String

'--- 매크로 ---------------------------------------------------
Public Sub SheetNav_Toggle()
    If mOn Then SheetNav_Off Else SheetNav_On False
End Sub

Public Sub SheetNav_On(ByVal quiet As Boolean)
    If mOn Then Exit Sub
    Set mApi = CoLibClient_Api(MIN_LIB, ADDIN_NAME, quiet)
    If mApi Is Nothing Then Exit Sub
    mRight = (mApi.CoSetting_Choice(ADDIN_ID, "side", SIDE_RIGHT, SIDE_RIGHT, SIDE_LEFT) = SIDE_RIGHT)
    mWidth = mApi.CoSetting_Number(ADDIN_ID, "width", 260, 160, 560)
    mFontPt = mApi.CoSetting_Number(ADDIN_ID, "fontSize", 13, 9, 22)
    mShowHidden = mApi.CoSetting_Bool(ADDIN_ID, "showHidden", False)
    Set mForm = New SheetNavForm
    mForm.Build mFontPt, mShowHidden
    Dock
    mForm.Show vbModeless
    mForm.Relayout
    On Error Resume Next
    AppActivate Application.Caption                    ' 키보드 입력이 엑셀로 돌아가게(창은 그대로 뜬 채)
    On Error GoTo 0
    mOn = True
    Set mEvents = New SheetNavEvents
    Set mEvents.App = Application
    SheetNav_Refresh
End Sub

Public Sub SheetNav_Off()
    mOn = False
    Set mEvents = Nothing
    If Not mForm Is Nothing Then
        On Error Resume Next
        Unload mForm
        On Error GoTo 0
        Set mForm = Nothing
    End If
    Set mApi = Nothing
End Sub

' 창의 X 단추로 닫았을 때(SheetNavForm에서 부른다)
Public Sub SheetNav_FormClosed()
    mOn = False
    Set mEvents = Nothing
    Set mForm = Nothing
    Set mApi = Nothing
End Sub

'--- 폼·이벤트에서 부르는 도우미 ---------------------------------------
' 활성 통합 문서의 시트 목록을 다시 만들어 창에 넣는다.
Public Sub SheetNav_Refresh()
    Dim wb As Workbook, sh As Object, names() As String, hid() As Boolean
    Dim n As Long, act As String
    If Not mOn Or mForm Is Nothing Then Exit Sub
    On Error Resume Next
    Set wb = ActiveWorkbook
    On Error GoTo EH
    If wb Is Nothing Then mForm.SetItems "", names, hid, 0, "": Exit Sub
    For Each sh In wb.Sheets
        If IsListed(sh.Visible, mShowHidden) Then
            n = n + 1
            ReDim Preserve names(1 To n): ReDim Preserve hid(1 To n)
            names(n) = sh.Name
            hid(n) = (sh.Visible <> xlSheetVisible)
        End If
    Next sh
    On Error Resume Next
    act = ActiveSheet.Name
    On Error GoTo EH
    mForm.SetItems wb.Name, names, hid, n, act
    Exit Sub
EH:
    ' 목록을 못 만들어도 엑셀 작업을 막지 않는다
End Sub

' 목록에서 시트 이름을 눌렀을 때 그 시트로 이동한다.
Public Sub SheetNav_Go(ByVal wbName As String, ByVal sheetName As String)
    Dim sh As Object
    On Error GoTo EH
    Set sh = Workbooks(wbName).Sheets(sheetName)
    If sh.Visible <> xlSheetVisible Then
        If Not mApi Is Nothing Then
            If Not mApi.CoGuard_Ask("'" & sheetName & "' 시트는 숨겨져 있습니다." & vbCrLf & "보이게 하고 이동할까요?", ADDIN_NAME) Then Exit Sub
        End If
        sh.Visible = xlSheetVisible
    End If
    Workbooks(wbName).Activate
    sh.Activate
    On Error Resume Next
    AppActivate Application.Caption                    ' 이동한 뒤 바로 키보드·마우스를 시트에서 쓰게
    Exit Sub
EH:
    Application.StatusBar = "시트 사이드바: '" & sheetName & "'(으)로 이동하지 못했습니다. 목록을 새로 고칩니다."
    SheetNav_Refresh
End Sub

Public Sub SheetNav_SwapSide()
    mRight = Not mRight
    Dock
End Sub

Public Sub SheetNav_SetShowHidden(ByVal v As Boolean)
    mShowHidden = v
    SheetNav_Refresh
End Sub

' 엑셀 창 가장자리에 맞춘다.
Public Sub SheetNav_Dock()
    If Not mOn Then Exit Sub
    Dock
End Sub

'--- 단축키·시작·종료 (ThisWorkbook 코드에서 부른다) ----------------
Public Sub SheetNav_Startup()
    Dim api As Object
    On Error Resume Next
    Set api = CoLibClient_Api(MIN_LIB, ADDIN_NAME, True)
    mKey = DEFAULT_KEY
    If Not api Is Nothing Then mKey = api.CoSetting_Text(ADDIN_ID, "key.toggle", DEFAULT_KEY)
    If Len(mKey) > 0 Then Application.OnKey mKey, "'" & ThisWorkbook.Name & "'!SheetNav_Toggle"
    If Not api Is Nothing Then
        If api.CoSetting_Bool(ADDIN_ID, "startOn", False) Then SheetNav_On True
    End If
End Sub

Public Sub SheetNav_Shutdown()
    On Error Resume Next
    SheetNav_Off
    If Len(mKey) > 0 Then Application.OnKey mKey
    mKey = ""
End Sub

'--- 이 추가기능만의 도우미 ---------------------------------------
Private Sub Dock()
    If mForm Is Nothing Then Exit Sub
    On Error Resume Next
    mForm.SetBounds DockX(Application.Left, Application.Width, mWidth, mRight), Application.Top + 150, mWidth, DockH(Application.Height)
End Sub

' 창을 붙일 왼쪽 위치. 오른쪽이면 엑셀 창 오른쪽 끝에서 너비만큼 안쪽.
Private Function DockX(ByVal appLeft As Double, ByVal appW As Double, ByVal w As Double, ByVal onRight As Boolean) As Double
    If onRight Then DockX = appLeft + appW - w - MARGIN_PT Else DockX = appLeft + MARGIN_PT
End Function

' 창 높이. 엑셀 창 높이에서 리본·수식 입력줄·상태 표시줄 몫을 뺀다(너무 작으면 최소 높이).
Private Function DockH(ByVal appH As Double) As Double
    DockH = appH - 220
    If DockH < 260 Then DockH = 260
End Function

' 목록에 넣을 시트인지(보이는 시트는 항상, 숨김은 설정이 켜졌을 때만, 아주 숨김은 절대 안 넣음).
Private Function IsListed(ByVal vis As Long, ByVal showHidden As Boolean) As Boolean
    If vis = xlSheetVisible Then
        IsListed = True
    ElseIf vis = xlSheetHidden Then
        IsListed = showHidden
    End If
End Function

'--- 시험 (직접 실행 창에서 실행) ----------------------------------
Public Sub Test_SheetNav()
    Dim f As Long
    f = f + Chk("오른쪽 붙이기", DockX(100, 1000, 200, True) = 100 + 1000 - 200 - MARGIN_PT)
    f = f + Chk("왼쪽 붙이기", DockX(100, 1000, 200, False) = 100 + MARGIN_PT)
    f = f + Chk("높이", DockH(800) = 580)
    f = f + Chk("최소 높이", DockH(300) = 260)
    f = f + Chk("보이는 시트", IsListed(xlSheetVisible, False))
    f = f + Chk("숨김 시트 끔", Not IsListed(xlSheetHidden, False))
    f = f + Chk("숨김 시트 켬", IsListed(xlSheetHidden, True))
    f = f + Chk("아주 숨김은 항상 제외", Not IsListed(xlSheetVeryHidden, True))
    f = f + Chk("시트 이름 걸러내기", SheetNav_Match("2026 견적서", "견적") And Not SheetNav_Match("요약", "견적") And SheetNav_Match("ABC", "bc") And SheetNav_Match("아무거나", ""))
    If f = 0 Then
        Debug.Print "SheetNav 시험: 모두 통과"
        MsgBox "SheetNav 시험: 모두 통과", vbInformation, "칼퇴"
    Else
        MsgBox "SheetNav 시험: " & f & "개 실패 (직접 실행 창 확인)", vbExclamation, "칼퇴"
    End If
End Sub

Private Function Chk(ByVal testName As String, ByVal ok As Boolean) As Long
    If Not ok Then Debug.Print "[실패] " & testName: Chk = 1
End Function

' 검색어로 시트 이름 걸러내기(대소문자 무시, 빈 검색어는 전부). 폼과 시험이 같이 쓴다.
Public Function SheetNav_Match(ByVal sheetName As String, ByVal query As String) As Boolean
    If Len(query) = 0 Then SheetNav_Match = True: Exit Function
    SheetNav_Match = (InStr(1, sheetName, query, vbTextCompare) > 0)
End Function
