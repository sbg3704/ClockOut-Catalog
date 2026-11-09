Attribute VB_Name = "SelStats"
'@clockout id=selection-stats name=선택 요약 크게 보기 host=excel
'@desc 선택한 칸의 합계·평균·개수를 상태 표시줄보다 훨씬 크게, 떠 있는 작은 창으로 보여 줍니다.
'==============================================================
' 기능: 엑셀 아래 상태 표시줄의 평균·개수·합계를 큰 글자로 보여 주는 창을 띄우고, 선택이 바뀔 때마다 갱신한다.
' 매크로:
'   SelStats_Toggle   창 열기/닫기 (기본 단축키 없음, 설정에서 지정)
' 사용법: 단축키나 매크로로 켜면 창이 뜨고, 칸을 고를 때마다 값이 바뀐다. 글자 크기·보일 값·쉼표·단축키는
'         칼퇴 설정창에서 바꾼다(선언: selection-stats.clockout.json, 값 저장: %APPDATA%\ClockOut\settings\selection-stats.json).
' 규칙: 값 계산은 SelStatsCalc.bas, 화면은 SelStatsForm.frm, 이벤트는 SelStatsEvents.cls에 있다.
'       숨긴 줄(필터로 가려진 줄 포함)과 오류 칸은 세지 않는다.
' 의존: 공용 라이브러리 clockout-lib 0.1.001 이상(CoSetting_*). CoLibClient.bas가 연결한다.
'==============================================================
Option Explicit

Private Const ADDIN_ID As String = "selection-stats"   ' .clockout.json의 id와 같게
Private Const ADDIN_NAME As String = "선택 요약 크게 보기"
Private Const MIN_LIB As String = "0.1.001"            ' .clockout.json의 dependencies.min_version과 같게
Private Const DEFAULT_KEY As String = ""

Private mForm As SelStatsForm
Private mEvents As SelStatsEvents
Private mOn As Boolean
Private mThousands As Boolean
Private mShowNum As Boolean
Private mShowMinMax As Boolean
Private mKey As String
Private mLeft As Double
Private mTop As Double

'--- 매크로 ---------------------------------------------------
Public Sub SelStats_Toggle()
    If mOn Then SelStats_Off Else SelStats_On False
End Sub

Public Sub SelStats_On(ByVal quiet As Boolean)
    Dim api As Object, big As Double, caps As Variant
    If mOn Then Exit Sub
    Set api = CoLibClient_Api(MIN_LIB, ADDIN_NAME, quiet)
    If api Is Nothing Then Exit Sub
    big = api.CoSetting_Number(ADDIN_ID, "fontSize", 28, 12, 80)
    mThousands = api.CoSetting_Bool(ADDIN_ID, "thousands", True)
    mShowNum = api.CoSetting_Bool(ADDIN_ID, "showNumCount", False)
    mShowMinMax = api.CoSetting_Bool(ADDIN_ID, "showMinMax", True)

    caps = Captions()
    Set mForm = New SelStatsForm
    mForm.Build caps, big
    PlaceForm
    mForm.Show vbModeless
    On Error Resume Next
    AppActivate Application.Caption                    ' 키보드 입력이 엑셀로 돌아가게(창은 그대로 뜬 채)
    On Error GoTo 0
    mOn = True
    Set mEvents = New SelStatsEvents
    Set mEvents.App = Application
    SelStats_Update
End Sub

Public Sub SelStats_Off()
    mOn = False
    Set mEvents = Nothing
    If Not mForm Is Nothing Then
        On Error Resume Next
        mLeft = mForm.Left: mTop = mForm.Top
        Unload mForm
        On Error GoTo 0
        Set mForm = Nothing
    End If
End Sub

' 창의 X 단추로 닫았을 때(SelStatsForm에서 부른다)
Public Sub SelStats_FormClosed()
    mOn = False
    Set mEvents = Nothing
    Set mForm = Nothing
End Sub

'--- 이벤트에서 부르는 도우미 ---------------------------------------
Public Sub SelStats_Update()
    Dim r As SelStatsResult, i As Long, rng As Range
    If Not mOn Or mForm Is Nothing Then Exit Sub
    On Error GoTo EH
    If TypeName(Application.Selection) <> "Range" Then SetAll "-": Exit Sub    ' 도형·차트를 골랐을 때
    Set rng = Application.Selection
    If Not SelStatsCalc_Run(rng, r) Then SetAll "-": Exit Sub
    i = 1
    mForm.SetValue i, Fmt(r.Sum, r.NumCount > 0): i = i + 1
    mForm.SetValue i, Fmt(r.Avg, r.NumCount > 0): i = i + 1
    mForm.SetValue i, CStr(r.Count): i = i + 1
    If mShowNum Then mForm.SetValue i, CStr(r.NumCount): i = i + 1
    If mShowMinMax Then
        mForm.SetValue i, Fmt(r.Min, r.NumCount > 0): i = i + 1
        mForm.SetValue i, Fmt(r.Max, r.NumCount > 0)
    End If
    Exit Sub
EH:
    SetAll "-"
End Sub

'--- 단축키·시작·종료 (ThisWorkbook 코드에서 부른다) ----------------
Public Sub SelStats_Startup()
    Dim api As Object
    On Error Resume Next
    Set api = CoLibClient_Api(MIN_LIB, ADDIN_NAME, True)
    mKey = DEFAULT_KEY
    If Not api Is Nothing Then mKey = api.CoSetting_Text(ADDIN_ID, "keyToggle", DEFAULT_KEY)
    If Len(mKey) > 0 Then Application.OnKey mKey, "'" & ThisWorkbook.Name & "'!SelStats_Toggle"
    If Not api Is Nothing Then
        If api.CoSetting_Bool(ADDIN_ID, "startOn", False) Then SelStats_On True
    End If
End Sub

Public Sub SelStats_Shutdown()
    On Error Resume Next
    SelStats_Off
    If Len(mKey) > 0 Then Application.OnKey mKey
    mKey = ""
End Sub

'--- 이 추가기능만의 도우미 ---------------------------------------
Private Function Captions() As Variant
    Dim a As Variant
    a = Array("합계", "평균", "개수")
    If mShowNum Then ReDim Preserve a(0 To UBound(a) + 1): a(UBound(a)) = "숫자 개수"
    If mShowMinMax Then
        ReDim Preserve a(0 To UBound(a) + 2)
        a(UBound(a) - 1) = "최소": a(UBound(a)) = "최대"
    End If
    Captions = a
End Function

Private Function Fmt(ByVal v As Double, ByVal hasNum As Boolean) As String
    If hasNum Then Fmt = SelStatsCalc_Fmt(v, mThousands) Else Fmt = "-"
End Function

Private Sub SetAll(ByVal text As String)
    Dim i As Long, n As Long
    n = 3 + IIf(mShowNum, 1, 0) + IIf(mShowMinMax, 2, 0)
    For i = 1 To n
        mForm.SetValue i, text
    Next i
End Sub

' 처음엔 엑셀 창 오른쪽 위, 닫았다 다시 열면 마지막 자리.
Private Sub PlaceForm()
    If mLeft > 0 Or mTop > 0 Then
        mForm.Left = mLeft: mForm.Top = mTop
    Else
        mForm.Left = Application.Left + Application.Width - mForm.Width - 40
        mForm.Top = Application.Top + 150
        If mForm.Left < 0 Then mForm.Left = 0
    End If
End Sub
