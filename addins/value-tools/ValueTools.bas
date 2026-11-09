Attribute VB_Name = "ValueTools"
'@clockout id=value-tools name=값 정리 도구 host=excel
'@desc 선택한 칸의 값을 한 번에 고치는 작은 도구 모음입니다. 공백 정리, 텍스트 숫자 변환, 병합 해제 후 값 채우기, 빈 칸 채우기, 올림·절사, 한글 금액, 두 열 비교.
'==============================================================
' 기능: 셀을 우클릭하면 나오는 "칼퇴 값 정리" 메뉴를 만들고, 각 도구가 같이 쓰는 준비·확인·마무리 도우미를 둔다.
' 메뉴(각 도구는 다른 모듈에 있다):
'   공백·숨은문자 정리   ValueFix_CleanSpaces       (ValueFix.bas)
'   텍스트 숫자 -> 숫자  ValueFix_TextToNumber       (ValueFix.bas)
'   병합 해제 + 값 채우기 ValueFix_UnmergeFill        (ValueFix.bas)
'   빈 칸 위 값으로 채우기 ValueFix_FillBlankDown     (ValueFix.bas)
'   올림·절사·반올림      ValueMath_RoundValues       (ValueMath.bas)
'   한글 금액(일금 ○○원정) ValueMath_KoreanWon        (ValueMath.bas)
'   두 열 다른 칸 선택    ValueCompare_SelectDiff     (ValueCompare.bas)
' 규칙: 값을 바꾸는 도구는 실행 전에 "N칸이 바뀝니다" 확인창을 띄운다. 수식 칸과 보호된 시트의 잠긴 칸은 건드리지 않는다.
'       실행 후에는 Ctrl+Z로 되돌릴 수 없다(VBA 한계). 설정(옵션)은 칼퇴 설정창(value-tools.clockout.json).
' 의존: 공용 라이브러리 clockout-lib 0.1.001 이상(CoRange_*, CoGuard_*, CoText_*, CoNum_*, CoSetting_*).
'       CoLibClient.bas가 연결한다.
'==============================================================
Option Explicit
Option Private Module

Public Const VT_ADDIN_ID As String = "value-tools"     ' .clockout.json의 id와 같게
Private Const ADDIN_NAME As String = "값 정리 도구"
Private Const MIN_LIB As String = "0.1.001"            ' .clockout.json의 dependencies.min_version과 같게
Private Const MENU_TAG As String = "ClockOutValueTools"

'--- 메뉴 (ThisWorkbook 코드에서 열 때·닫을 때 부른다) -------------------
Public Sub ValueTools_Startup()
    On Error Resume Next
    ValueTools_Shutdown
    Dim pop As Object, items As Variant, i As Long
    items = Array("공백·숨은문자 정리|ValueFix_CleanSpaces", _
                  "텍스트 숫자를 숫자로|ValueFix_TextToNumber", _
                  "병합 해제 + 값 채우기|ValueFix_UnmergeFill", _
                  "빈 칸을 위 값으로 채우기|ValueFix_FillBlankDown", _
                  "올림·절사·반올림(실제 값)|ValueMath_RoundValues", _
                  "한글 금액(일금 ○○원정)|ValueMath_KoreanWon", _
                  "두 열 다른 칸 선택|ValueCompare_SelectDiff")
    Set pop = Application.CommandBars("Cell").Controls.Add(Type:=10, Temporary:=True)   ' 10 = 팝업 메뉴
    pop.Caption = "칼퇴 값 정리"
    pop.Tag = MENU_TAG
    For i = LBound(items) To UBound(items)
        With pop.Controls.Add(Type:=1, Temporary:=True)                              ' 1 = 단추
            .Caption = Split(items(i), "|")(0)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & Split(items(i), "|")(1)
        End With
    Next i
End Sub

Public Sub ValueTools_Shutdown()
    Dim i As Long, ctl As Object
    On Error Resume Next
    For i = Application.CommandBars("Cell").Controls.Count To 1 Step -1
        Set ctl = Application.CommandBars("Cell").Controls(i)
        If ctl.Tag = MENU_TAG Then ctl.Delete
    Next i
End Sub

'--- 도구들이 같이 쓰는 도우미 ---------------------------------------
' 공용 라이브러리 개체(없거나 판이 낮으면 안내 후 Nothing).
Public Function VT_Api() As Object
    Set VT_Api = CoLibClient_Api(MIN_LIB, ADDIN_NAME)
End Function

' 준비: 라이브러리 개체, 활성 시트, 선택한 칸 덩어리들(사용 중인 범위와 겹친 부분)을 얻는다. 안 되면 False(안내 끝남).
Public Function VT_Prepare(ByRef api As Object, ByRef ws As Worksheet, ByRef areas As Collection, _
                           Optional ByVal needEdit As Boolean = True) As Boolean
    Dim why As String, sel As Range, ar As Range, part As Range
    Set api = VT_Api()
    If api Is Nothing Then Exit Function
    If Not api.CoRange_GetSelection(ws, sel, why, True, True) Then api.CoGuard_Notice why: Exit Function
    If needEdit Then
        If Not api.CoRange_CanEdit(sel) Then
            api.CoGuard_Notice "시트가 보호되어 있고 잠긴 칸이 포함되어 있어 값을 바꿀 수 없습니다." & vbCrLf & _
                               "검토 > 시트 보호 해제 후 다시 실행하세요."
            Exit Function
        End If
    End If
    Set areas = New Collection
    For Each ar In sel.Areas
        Set part = Intersect(ar, ws.UsedRange)           ' 열·행 전체를 골라도 사용 중인 부분만 돈다
        If Not part Is Nothing Then areas.Add part
    Next ar
    If areas.Count = 0 Then api.CoGuard_Notice "선택한 칸에 값이 없습니다.": Exit Function
    VT_Prepare = True
End Function

' 값을 바꾸기 전 확인. 되돌릴 수 없음을 함께 알린다.
Public Function VT_Confirm(ByVal api As Object, ByVal what As String) As Boolean
    VT_Confirm = api.CoGuard_Ask(what & vbCrLf & vbCrLf & "실행 후에는 Ctrl+Z로 되돌릴 수 없습니다. 진행할까요?")
End Function

' 범위의 수식(상수는 값 그대로)을 항상 2차원 배열로. 한 칸이어도 (1,1) 배열.
Public Function VT_Read(ByVal rng As Range) As Variant
    Dim v As Variant, one(1 To 1, 1 To 1) As Variant
    v = rng.Formula
    If IsArray(v) Then
        VT_Read = v
    Else
        one(1, 1) = v
        VT_Read = one
    End If
End Function

' 범위의 값(Value2)을 항상 2차원 배열로. 한 칸이어도 (1,1) 배열.
Public Function VT_ReadValues(ByVal rng As Range) As Variant
    Dim v As Variant, one(1 To 1, 1 To 1) As Variant
    v = rng.Value2
    If IsArray(v) Then
        VT_ReadValues = v
    Else
        one(1, 1) = v
        VT_ReadValues = one
    End If
End Function

' 수식이 아닌 글자(상수 문자열) 칸이면 True.
Public Function VT_IsConstText(ByVal v As Variant) As Boolean
    If VarType(v) <> vbString Then Exit Function
    If Len(v) = 0 Then Exit Function
    VT_IsConstText = (Left$(v, 1) <> "=")
End Function

' 글자를 그대로 글자로 쓴다. 숫자처럼 보이는 글자가 숫자로 바뀌지 않게, 서식이 일반이면 앞에 '를 붙인다.
Public Sub VT_SetText(ByVal cel As Range, ByVal text As String)
    If cel.NumberFormat = "@" Then
        cel.Value2 = text
    Else
        cel.Value2 = "'" & text
    End If
End Sub

' 숫자를 쓴다. 서식이 텍스트(@)면 숫자로 읽히도록 일반으로 바꾼 뒤 쓴다.
Public Sub VT_SetNumber(ByVal cel As Range, ByVal num As Double)
    If cel.NumberFormat = "@" Then cel.NumberFormat = "General"
    cel.Value2 = num
End Sub

' 끝 안내(상태 표시줄 + 창).
Public Sub VT_Done(ByVal api As Object, ByVal msg As String)
    api.CoGuard_Notice msg
End Sub
