Attribute VB_Name = "QuickTools"
'@clockout id=quick-tools name=빠른 셀 도구 host=excel
'@desc 셀 우클릭 "칼퇴 빠른 도구" 메뉴. 값으로 고정, IFERROR로 감싸기, 부호 뒤집기, 앞뒤 글자 붙이기, 병합을 선택 영역 가운데로, 텍스트 날짜를 날짜로, 조건 선택.
'==============================================================
' 기능: 셀 우클릭 메뉴("칼퇴 빠른 도구")를 만들고 도구들이 같이 쓰는 도우미를 둔다.
' 메뉴(각 도구는 다른 모듈에 있다):
'   값으로 고정            QuickEdit_ToValues       (QuickEdit.bas)
'   IFERROR로 감싸기       QuickEdit_WrapIferror    (QuickEdit.bas)
'   부호 뒤집기            QuickEdit_FlipSign       (QuickEdit.bas)
'   앞뒤 글자 붙이기       QuickEdit_AddAffix       (QuickEdit.bas)
'   병합 -> 선택 영역 가운데  QuickMisc_MergeToCenter (QuickMisc.bas)
'   텍스트 날짜를 날짜로    QuickMisc_TextToDate     (QuickMisc.bas)
'   조건으로 칸 선택        QuickMisc_SelectSpecial  (QuickMisc.bas)
' 규칙: 값을 바꾸는 도구는 "N칸이 바뀝니다" 확인 후 실행. 실행 후 Ctrl+Z 불가. 보호된 시트의 잠긴 칸은 하지 않는다.
' 의존: 공용 라이브러리 clockout-lib 0.1.001 이상. CoLibClient.bas가 연결한다.
'==============================================================
Option Explicit
Option Private Module

Public Const QT_ADDIN_ID As String = "quick-tools"
Private Const ADDIN_NAME As String = "빠른 셀 도구"
Private Const MIN_LIB As String = "0.1.001"
Private Const MENU_TAG As String = "ClockOutQuickTools"

Public Sub QuickTools_Startup()
    On Error Resume Next
    QuickTools_Shutdown
    Dim pop As Object, items As Variant, i As Long
    items = Array("값으로 고정|QuickEdit_ToValues", _
                  "IFERROR로 감싸기|QuickEdit_WrapIferror", _
                  "부호 뒤집기(x -1)|QuickEdit_FlipSign", _
                  "앞뒤 글자 붙이기|QuickEdit_AddAffix", _
                  "병합 해제 -> 선택 영역 가운데|QuickMisc_MergeToCenter", _
                  "텍스트 날짜를 날짜로|QuickMisc_TextToDate", _
                  "조건으로 칸 선택(병합·오류·수식…)|QuickMisc_SelectSpecial")
    Set pop = Application.CommandBars("Cell").Controls.Add(Type:=10, Temporary:=True)
    pop.Caption = "칼퇴 빠른 도구"
    pop.Tag = MENU_TAG
    For i = LBound(items) To UBound(items)
        With pop.Controls.Add(Type:=1, Temporary:=True)
            .Caption = Split(items(i), "|")(0)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & Split(items(i), "|")(1)
        End With
    Next i
End Sub

Public Sub QuickTools_Shutdown()
    Dim i As Long, ctl As Object
    On Error Resume Next
    For i = Application.CommandBars("Cell").Controls.Count To 1 Step -1
        Set ctl = Application.CommandBars("Cell").Controls(i)
        If ctl.Tag = MENU_TAG Then ctl.Delete
    Next i
End Sub

Public Function QT_Api() As Object
    Set QT_Api = CoLibClient_Api(MIN_LIB, ADDIN_NAME)
End Function

' 준비: 라이브러리 개체, 활성 시트, 선택한 칸 덩어리들(사용 중인 범위와 겹친 부분). 안 되면 False(안내 끝남).
Public Function QT_Prepare(ByRef api As Object, ByRef ws As Worksheet, ByRef areas As Collection, Optional ByVal needEdit As Boolean = True) As Boolean
    Dim why As String, sel As Range, ar As Range, part As Range
    Set api = QT_Api()
    If api Is Nothing Then Exit Function
    If Not api.CoRange_GetSelection(ws, sel, why, True, True) Then api.CoGuard_Notice why: Exit Function
    If needEdit Then
        If Not api.CoRange_CanEdit(sel) Then
            api.CoGuard_Notice "시트가 보호되어 있고 잠긴 칸이 포함되어 있어 값을 바꿀 수 없습니다." & vbCrLf & "검토 > 시트 보호 해제 후 다시 실행하세요."
            Exit Function
        End If
    End If
    Set areas = New Collection
    For Each ar In sel.Areas
        Set part = Intersect(ar, ws.UsedRange)
        If Not part Is Nothing Then areas.Add part
    Next ar
    If areas.Count = 0 Then api.CoGuard_Notice "선택한 칸에 값이 없습니다.": Exit Function
    QT_Prepare = True
End Function

Public Function QT_Confirm(ByVal api As Object, ByVal what As String) As Boolean
    QT_Confirm = api.CoGuard_Ask(what & vbCrLf & vbCrLf & "실행 후에는 Ctrl+Z로 되돌릴 수 없습니다. 진행할까요?")
End Function

' 범위의 수식(상수는 값 그대로)을 항상 2차원 배열로.
Public Function QT_Read(ByVal rng As Range) As Variant
    Dim v As Variant, one(1 To 1, 1 To 1) As Variant
    v = rng.Formula
    If IsArray(v) Then
        QT_Read = v
    Else
        one(1, 1) = v
        QT_Read = one
    End If
End Function
