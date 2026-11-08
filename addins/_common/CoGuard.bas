Attribute VB_Name = "CoGuard"
'==============================================================
' 기능: 칼퇴 공용 라이브러리 - 실행 상태 가드와 안내 창.
'       매크로가 ScreenUpdating·EnableEvents·계산 방식·상태 표시줄·커서를 바꿨다가
'       오류가 나도 항상 원래대로 돌려놓게 한다(image-fit의 prevUpdating 패턴을 일반화).
' 사용법:
'   Dim st As CoAppState, msg As String
'   st = CoGuard_Begin()                     ' 화면 갱신 끔(기본). 계산 수동·이벤트 끔은 인자로
'   On Error GoTo EH
'   ... 작업 ...
'   CoGuard_End st
'   Exit Sub
' EH:
'   msg = CoGuard_ErrText("사진을 넣는 중")  ' 반드시 맨 먼저(다음 줄의 On Error가 Err를 지운다)
'   CoGuard_End st
'   CoGuard_Notice msg
' 의존: 없음(Excel Application만). Option Private Module이라 시트 함수·매크로 목록에 안 보인다.
'==============================================================
Option Explicit
Option Private Module

Public Const CO_MSG_TITLE As String = "칼퇴"

Public Type CoAppState
    Saved As Boolean
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    DisplayAlerts As Boolean
    CalcSaved As Boolean
    Calculation As Long
    StatusBar As Variant
    Cursor As Long
End Type

' 지금 상태를 저장하고 필요한 것만 끈다. 돌려받은 값을 CoGuard_End에 그대로 넘긴다.
Public Function CoGuard_Begin(Optional ByVal noScreen As Boolean = True, _
                              Optional ByVal noEvents As Boolean = False, _
                              Optional ByVal manualCalc As Boolean = False, _
                              Optional ByVal noAlerts As Boolean = False) As CoAppState
    Dim s As CoAppState
    On Error Resume Next
    s.ScreenUpdating = Application.ScreenUpdating
    s.EnableEvents = Application.EnableEvents
    s.DisplayAlerts = Application.DisplayAlerts
    s.StatusBar = Application.StatusBar
    s.Cursor = Application.Cursor
    Err.Clear
    s.Calculation = Application.Calculation          ' 열린 통합 문서가 없으면 오류 -> 저장 안 함
    s.CalcSaved = (Err.Number = 0)
    Err.Clear
    s.Saved = True

    If noScreen Then Application.ScreenUpdating = False
    If noEvents Then Application.EnableEvents = False
    If noAlerts Then Application.DisplayAlerts = False
    If manualCalc And s.CalcSaved Then Application.Calculation = xlCalculationManual
    CoGuard_Begin = s
End Function

' 저장한 상태로 되돌린다. 두 번 불러도 안전하다(두 번째는 아무것도 안 함).
Public Sub CoGuard_End(ByRef s As CoAppState)
    If Not s.Saved Then Exit Sub
    On Error Resume Next
    If s.CalcSaved Then
        If Application.Calculation <> s.Calculation Then Application.Calculation = s.Calculation
    End If
    Application.EnableEvents = s.EnableEvents
    Application.DisplayAlerts = s.DisplayAlerts
    Application.StatusBar = s.StatusBar              ' False면 Excel 기본 표시로 돌아간다
    Application.Cursor = s.Cursor
    Application.ScreenUpdating = s.ScreenUpdating
    s.Saved = False
End Sub

' 상태 표시줄에 진행 상황을 보인다. CoGuard_End가 원래대로 돌려놓는다.
Public Sub CoGuard_Status(ByVal text As String)
    On Error Resume Next
    Application.StatusBar = text
End Sub

' 오류 안내 글. 오류 처리부 맨 처음에 불러야 Err 내용이 남아 있다(이 함수는 Err를 지우지 않는다).
Public Function CoGuard_ErrText(Optional ByVal context As String = "작업 중") As String
    If Err.Number = 0 Then
        CoGuard_ErrText = context & " 알 수 없는 문제가 생겼습니다."
    Else
        CoGuard_ErrText = context & " 오류가 났습니다." & vbCrLf & _
                          Err.Description & " (오류 " & Err.Number & ")"
    End If
End Function

Public Sub CoGuard_Notice(ByVal msg As String, Optional ByVal title As String = CO_MSG_TITLE)
    MsgBox msg, vbExclamation, title
End Sub

Public Sub CoGuard_Info(ByVal msg As String, Optional ByVal title As String = CO_MSG_TITLE)
    MsgBox msg, vbInformation, title
End Sub

' 예/아니오를 묻는다. 기본 단추는 "아니오"(실수로 Enter를 눌러도 진행하지 않게).
Public Function CoGuard_Ask(ByVal msg As String, Optional ByVal title As String = CO_MSG_TITLE) As Boolean
    CoGuard_Ask = (MsgBox(msg, vbQuestion + vbYesNo + vbDefaultButton2, title) = vbYes)
End Function
