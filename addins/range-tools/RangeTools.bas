Attribute VB_Name = "RangeTools"
'@clockout id=range-tools name=범위 도구 host=excel
'@desc 셀 우클릭 "칼퇴 범위 도구" 메뉴. 자연 정렬(1호,2호,10호 순서), 색깔별 합계, 그룹별로 시트 쪼개기(소계 포함).
'==============================================================
' 기능: 셀 우클릭 메뉴("칼퇴 범위 도구")를 만들고 도구들이 같이 쓰는 도우미를 둔다.
' 메뉴(각 도구는 다른 모듈에 있다):
'   자연 정렬               RangeSort_Natural      (RangeSort.bas)
'   색깔별 합계             RangeColor_Sum         (RangeColor.bas)
'   그룹별 시트로 쪼개기    RangeSplit_ByGroup     (RangeSplit.bas)
' 의존: 공용 라이브러리 clockout-lib 0.1.001 이상. CoLibClient.bas가 연결한다.
'==============================================================
Option Explicit
Option Private Module

Public Const RT_ADDIN_ID As String = "range-tools"
Private Const ADDIN_NAME As String = "범위 도구"
Private Const MIN_LIB As String = "0.1.001"
Private Const MENU_TAG As String = "ClockOutRangeTools"

Public Sub RangeTools_Startup()
    On Error Resume Next
    RangeTools_Shutdown
    Dim pop As Object, items As Variant, i As Long
    items = Array("자연 정렬(1호, 2호, 10호 순서)|RangeSort_Natural", _
                  "색깔별 합계|RangeColor_Sum", _
                  "그룹별 시트로 쪼개기(소계)|RangeSplit_ByGroup")
    Set pop = Application.CommandBars("Cell").Controls.Add(Type:=10, Temporary:=True)
    pop.Caption = "칼퇴 범위 도구"
    pop.Tag = MENU_TAG
    For i = LBound(items) To UBound(items)
        With pop.Controls.Add(Type:=1, Temporary:=True)
            .Caption = Split(items(i), "|")(0)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & Split(items(i), "|")(1)
        End With
    Next i
End Sub

Public Sub RangeTools_Shutdown()
    Dim i As Long, ctl As Object
    On Error Resume Next
    For i = Application.CommandBars("Cell").Controls.Count To 1 Step -1
        Set ctl = Application.CommandBars("Cell").Controls(i)
        If ctl.Tag = MENU_TAG Then ctl.Delete
    Next i
End Sub

Public Function RT_Api() As Object
    Set RT_Api = CoLibClient_Api(MIN_LIB, ADDIN_NAME)
End Function

Public Function RT_Confirm(ByVal api As Object, ByVal what As String) As Boolean
    RT_Confirm = api.CoGuard_Ask(what & vbCrLf & vbCrLf & "실행 후에는 Ctrl+Z로 되돌릴 수 없습니다. 진행할까요?")
End Function

' 선택(한 덩어리)을 사용 중인 부분으로 줄여 돌려준다. 안 되면 안내 후 Nothing.
Public Function RT_Area(ByVal api As Object, ByRef ws As Worksheet) As Range
    Dim sel As Range, why As String, part As Range
    If Not api.CoRange_GetSelection(ws, sel, why, False, True) Then api.CoGuard_Notice why: Exit Function
    If sel.Cells.CountLarge = 1 Then Set sel = sel.CurrentRegion          ' 한 칸이면 이어진 표 전체
    Set part = Intersect(sel, ws.UsedRange)
    If part Is Nothing Then api.CoGuard_Notice "선택한 칸에 값이 없습니다.": Exit Function
    Set RT_Area = part
End Function
