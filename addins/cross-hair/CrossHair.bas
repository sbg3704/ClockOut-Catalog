Attribute VB_Name = "CrossHair"
'@clockout id=cross-hair name=십자 음영 host=excel
'@desc 선택한 칸의 행과 열 전체에 음영을 넣어 큰 화면에서도 지금 어느 칸인지 한눈에 보이게 합니다.
'==============================================================
' 기능: 활성 셀과 같은 행·같은 열을 눈에 띄게 한다. 기본은 행의 위·아래 선과 열의 좌·우 선(칸의 배경색을
'       덮지 않음)이고, 설정으로 연한 음영을 더하거나 음영만 쓸 수도 있다. 셀 값은 건드리지 않고, 시트 전체에
'       표시 전용 조건부 서식 규칙을 걸어 둔다. 끄거나 저장할 때 규칙을 지워 파일에 흔적이 남지 않는다.
' 매크로:
'   CrossHair_Toggle   켜기/끄기 (기본 단축키 없음, 설정에서 지정)
' 사용법: 단축키나 매크로로 켜면 선택을 옮길 때마다 따라온다. 색·행/열 선택·단축키는 칼퇴 설정창에서 바꾼다
'         (선언: cross-hair.clockout.json, 값 저장: %APPDATA%\ClockOut\settings\cross-hair.json).
' 원리: 규칙 수식이 CELL("row")·CELL("col")(지금 선택한 칸)을 읽고, 선택이 바뀔 때마다 그 시트만 다시
'       계산시켜 규칙이 새 위치를 읽게 한다. 서식 규칙이라 셀 값·수식은 바뀌지 않는다.
' 한계: 음영(채우기)은 엑셀 조건부 서식이라 투명하지 않아 칸의 원래 배경색을 덮는다. 그래서 기본은 아주 연한 음영 + 선이고, 선택한 칸만 진한 색이다.
'       보호된 시트에는 걸 수 없다.
' 의존: 공용 라이브러리 clockout-lib 0.1.001 이상(CoSetting_*, CoGuard_Notice). CoLibClient.bas가 연결한다.
'==============================================================
Option Explicit

Private Const ADDIN_ID As String = "cross-hair"     ' .clockout.json의 id와 같게
Private Const ADDIN_NAME As String = "십자 음영"
Private Const MIN_LIB As String = "0.1.001"          ' .clockout.json의 dependencies.min_version과 같게
Private Const MARK As String = "ClockOutCross"      ' 규칙 수식 안의 표식: 우리가 만든 규칙을 찾는 데 쓴다
Private Const DEFAULT_KEY As String = ""
Private Const DEFAULT_LINE As String = "#C5E4A8"
Private Const DEFAULT_FILL As String = "#F0F9E6"
Private Const DEFAULT_ACTIVE As String = "#B7E58E"
Private Const STYLE_LINE As String = "선"
Private Const STYLE_BOTH As String = "선 + 연한 음영"
Private Const STYLE_FILL As String = "음영만"

Private mEvents As CrossHairEvents
Private mOn As Boolean
Private mLineColor As Long
Private mFillColor As Long
Private mActiveColor As Long
Private mActiveOn As Boolean
Private mUseLine As Boolean
Private mUseFill As Boolean
Private mRowOn As Boolean
Private mColOn As Boolean
Private mKey As String

'--- 매크로 ---------------------------------------------------
Public Sub CrossHair_Toggle()
    If mOn Then CrossHair_Off Else CrossHair_On False
End Sub

Public Sub CrossHair_On(ByVal quiet As Boolean)
    Dim api As Object, st As String
    Set api = CoLibClient_Api(MIN_LIB, ADDIN_NAME, quiet)
    If api Is Nothing Then Exit Sub
    st = api.CoSetting_Choice(ADDIN_ID, "style", STYLE_BOTH, STYLE_LINE, STYLE_BOTH, STYLE_FILL)
    mUseLine = (st <> STYLE_FILL)
    mUseFill = (st <> STYLE_LINE)
    mLineColor = ParseHexColor(api.CoSetting_Text(ADDIN_ID, "lineColor", DEFAULT_LINE), RGB(197, 228, 168))
    mFillColor = ParseHexColor(api.CoSetting_Text(ADDIN_ID, "fillColor", DEFAULT_FILL), RGB(240, 249, 230))
    mActiveColor = ParseHexColor(api.CoSetting_Text(ADDIN_ID, "activeColor", DEFAULT_ACTIVE), RGB(183, 229, 142))
    mActiveOn = api.CoSetting_Bool(ADDIN_ID, "activeFill", True)
    mRowOn = api.CoSetting_Bool(ADDIN_ID, "showRow", True)
    mColOn = api.CoSetting_Bool(ADDIN_ID, "showCol", True)
    If Not mRowOn And Not mColOn Then mRowOn = True: mColOn = True     ' 둘 다 끄면 의미가 없어 둘 다 켠다
    StripAll                                           ' 설정이 바뀌었을 수 있으니 옛 규칙을 지우고 새로 건다
    mOn = True
    If mEvents Is Nothing Then Set mEvents = New CrossHairEvents
    Set mEvents.App = Application
    If Not ActiveSheet Is Nothing Then CrossHair_Apply ActiveSheet
End Sub

Public Sub CrossHair_Off()
    mOn = False
    Set mEvents = Nothing
    StripAll
End Sub

'--- 이벤트에서 부르는 도우미 ---------------------------------------
' 시트에 규칙이 없으면 걸고, 그 시트를 다시 계산시켜 규칙이 지금 선택한 칸을 읽게 한다.
Public Sub CrossHair_Apply(ByVal sh As Object)
    Dim ws As Worksheet
    If Not mOn Then Exit Sub
    If Not TypeOf sh Is Worksheet Then Exit Sub
    Set ws = sh
    On Error Resume Next                               ' 보호된 시트·특수 시트는 조용히 건너뛴다
    If ws.ProtectContents Then Exit Sub
    If Not HasRule(ws) Then AddRule ws
    ws.Calculate
End Sub

' 저장·닫기 직전에 규칙을 지운다(파일에 남기지 않기 위해).
Public Sub CrossHair_Strip(ByVal wb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In wb.Worksheets
        StripSheet ws
    Next ws
End Sub

' 저장 뒤에 다시 거는 일(OnTime으로 불린다).
Public Sub CrossHair_Reapply()
    On Error Resume Next
    If mOn And Not ActiveSheet Is Nothing Then CrossHair_Apply ActiveSheet
End Sub

'--- 단축키·시작·종료 (ThisWorkbook 코드에서 부른다) ----------------
Public Sub CrossHair_Startup()
    Dim api As Object
    On Error Resume Next
    Set api = CoLibClient_Api(MIN_LIB, ADDIN_NAME, True)
    mKey = DEFAULT_KEY
    If Not api Is Nothing Then mKey = api.CoSetting_Text(ADDIN_ID, "key.toggle", DEFAULT_KEY)
    If Len(mKey) > 0 Then Application.OnKey mKey, "'" & ThisWorkbook.Name & "'!CrossHair_Toggle"
    If Not api Is Nothing Then
        If api.CoSetting_Bool(ADDIN_ID, "startOn", False) Then CrossHair_On True
    End If
End Sub

Public Sub CrossHair_Shutdown()
    On Error Resume Next
    CrossHair_Off
    If Len(mKey) > 0 Then Application.OnKey mKey
    mKey = ""
End Sub

'--- 이 추가기능만의 도우미 ---------------------------------------
' 규칙 수식. which: "row"(행만), "col"(열만), "both"(행 또는 열).
Private Function RuleFormula(ByVal which As String) As String
    Dim parts As String
    If which = "row" Then
        parts = "ROW()=CELL(""row"")"
    ElseIf which = "col" Then
        parts = "COLUMN()=CELL(""col"")"
    ElseIf which = "cell" Then
        parts = "ROW()=CELL(""row""),COLUMN()=CELL(""col"")"
    Else
        parts = "OR(ROW()=CELL(""row""),COLUMN()=CELL(""col""))"
    End If
    RuleFormula = "=AND(LEN(""" & MARK & """)>0," & parts & ")"
End Function

Private Sub AddRule(ByVal ws As Worksheet)
    Dim which As String
    ' 먼저 건 규칙이 아래로 가므로 음영 -> 선 순서로 건다(선이 항상 위에 보이게 위로 올린다).
    If mUseFill Then
        If mRowOn And mColOn Then
            which = "both"
        ElseIf mRowOn Then
            which = "row"
        Else
            which = "col"
        End If
        With ws.Cells.FormatConditions.Add(Type:=xlExpression, Formula1:=RuleFormula(which))
            .SetFirstPriority
            .StopIfTrue = False                        ' 사용자가 만든 다른 규칙(글자색 등)도 그대로 보이게
            .Interior.Color = mFillColor
        End With
    End If
    If mUseLine Then
        If mRowOn Then AddLineRule ws, "row", xlTop, xlBottom
        If mColOn Then AddLineRule ws, "col", xlLeft, xlRight
    End If
    If mActiveOn Then                                  ' 선택한 칸 자신은 보조 줄보다 진한 색(맨 위 규칙)
        With ws.Cells.FormatConditions.Add(Type:=xlExpression, Formula1:=RuleFormula("cell"))
            .SetFirstPriority
            .StopIfTrue = False
            .Interior.Color = mActiveColor
        End With
    End If
End Sub

Private Sub AddLineRule(ByVal ws As Worksheet, ByVal which As String, ByVal side1 As Long, ByVal side2 As Long)
    With ws.Cells.FormatConditions.Add(Type:=xlExpression, Formula1:=RuleFormula(which))
        .SetFirstPriority
        .StopIfTrue = False
        With .Borders(side1)
            .LineStyle = xlContinuous
            .Color = mLineColor
        End With
        With .Borders(side2)
            .LineStyle = xlContinuous
            .Color = mLineColor
        End With
    End With
End Sub

Private Function HasRule(ByVal ws As Worksheet) As Boolean
    Dim i As Long
    On Error Resume Next
    For i = 1 To ws.Cells.FormatConditions.Count
        If InStr(ws.Cells.FormatConditions(i).Formula1, MARK) > 0 Then HasRule = True: Exit Function
    Next i
End Function

Private Sub StripSheet(ByVal ws As Worksheet)
    Dim i As Long
    On Error Resume Next
    If ws.ProtectContents Then Exit Sub
    For i = ws.Cells.FormatConditions.Count To 1 Step -1
        If InStr(ws.Cells.FormatConditions(i).Formula1, MARK) > 0 Then ws.Cells.FormatConditions(i).Delete
    Next i
End Sub

Private Sub StripAll()
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        CrossHair_Strip wb
    Next wb
End Sub

' "#RRGGBB" -> 색 값. 못 읽으면 fallback.
Private Function ParseHexColor(ByVal hex As String, ByVal fallback As Long) As Long
    Dim s As String
    s = Replace(Trim$(hex), "#", "")
    If Len(s) <> 6 Then ParseHexColor = fallback: Exit Function
    On Error GoTo Bad
    ParseHexColor = RGB(CLng("&H" & Mid$(s, 1, 2)), CLng("&H" & Mid$(s, 3, 2)), CLng("&H" & Mid$(s, 5, 2)))
    Exit Function
Bad:
    ParseHexColor = fallback
End Function

'--- 시험 (직접 실행 창에서 실행) ----------------------------------
Public Sub Test_CrossHair()
    Dim f As Long
    f = f + Chk("색 코드", ParseHexColor("#FF0000", 0) = RGB(255, 0, 0))
    f = f + Chk("색 코드 # 없이", ParseHexColor("00FF00", 0) = RGB(0, 255, 0))
    f = f + Chk("색 코드 잘못", ParseHexColor("빨강", 7) = 7)
    f = f + Chk("색 코드 자리수", ParseHexColor("#FFF", 9) = 9)
    f = f + Chk("행+열 수식", InStr(RuleFormula("both"), "OR(ROW()=CELL(""row""),COLUMN()=CELL(""col""))") > 0)
    f = f + Chk("행만 수식", InStr(RuleFormula("row"), "ROW()=CELL(""row"")") > 0 And InStr(RuleFormula("row"), "COLUMN") = 0)
    f = f + Chk("열만 수식", InStr(RuleFormula("col"), "COLUMN()=CELL(""col"")") > 0 And InStr(RuleFormula("col"), "ROW") = 0)
    f = f + Chk("선택한 칸 수식", InStr(RuleFormula("cell"), "AND(LEN(""" & MARK & """)>0,ROW()=CELL(""row""),COLUMN()=CELL(""col""))") > 0)
    f = f + Chk("표식", InStr(RuleFormula("row"), MARK) > 0)
    If f = 0 Then
        Debug.Print "CrossHair 시험: 모두 통과"
        MsgBox "CrossHair 시험: 모두 통과", vbInformation, "칼퇴"
    Else
        MsgBox "CrossHair 시험: " & f & "개 실패 (직접 실행 창 확인)", vbExclamation, "칼퇴"
    End If
End Sub

Private Function Chk(ByVal testName As String, ByVal ok As Boolean) As Long
    If Not ok Then Debug.Print "[실패] " & testName: Chk = 1
End Function
