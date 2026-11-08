Attribute VB_Name = "ImageFit"
'@clockout id=image-fit name=사진 셀에 맞추기 host=excel
'@desc 선택한 셀(병합 셀 포함)에 사진을 맞춰 넣습니다. 사진 파일 고르기, 클립보드 사진 붙여넣고 맞추기, 이미 올려 둔 사진 맞추기 3가지 기능이 있습니다.
'==============================================================
' 기능: 사진을 셀(또는 병합 셀, 선택 영역)에 맞춰 넣는다. 원본은 "사진매크로 시트.xlsm"의
'       insert_Pic(사진 파일 넣기)과 개체를_셀에맞춤(붙여넣기 후 맞춤)이다.
' 매크로:
'   ImageFit_InsertPicture  사진 파일을 골라 선택한 칸에 맞춰 넣기   (기본 단축키 Ctrl+Shift+A)
'   ImageFit_PasteAndFit    클립보드의 사진을 붙여 선택한 칸에 맞추기 (기본 단축키 Ctrl+Shift+V)
'   ImageFit_FitSelected    이미 올려 둔 사진(들)을 지금 놓인 칸에 맞추기
' 사용법: 칸을 선택하고 매크로를 실행한다. 여백·비율·정렬·단축키는 칼퇴 설정창에서 바꾼다
'         (선언: image-fit.clockout.json, 값 저장: %APPDATA%\ClockOut\settings\image-fit.json).
'         실행 후에는 Ctrl+Z로 되돌릴 수 없다(VBA 한계). 잘못 넣었으면 사진을 지우면 된다.
' 주의: 맨 위 '@clockout·'@desc 줄은 매니저가 코드에서 이름·설명(? 툴팁)을 읽는 자리다(docs/03 주석 블록 방식).
'       image-fit.clockout.json이 있으면 JSON이 우선이므로, 설명을 고칠 때 두 곳을 같이 맞춘다.
' 의존: 공용 라이브러리 addins/_common (ClockOutSettings, CoNum, CoSetting, CoGuard, CoRange,
'       CoFile, CoText, CoFitMath, CoShape). 공용 기능은 여기서 다시 만들지 않고 라이브러리를 쓴다.
'==============================================================
Option Explicit

Private Const ADDIN_ID As String = "image-fit"      ' .clockout.json의 id와 같게

' 설정 선택지 글자는 image-fit.clockout.json의 options와 정확히 같아야 한다.
Private Const OPT_ALIGN_CENTER As String = "가운데"
Private Const OPT_ALIGN_TOPLEFT As String = "왼쪽 위"
Private Const OPT_AREA_ALL As String = "선택 영역 전체"
Private Const OPT_AREA_FIRSTCELL As String = "선택한 첫 칸(병합 포함)"
Private Const OPT_PLACE_ALL As String = "셀과 함께 이동 및 크기 변경"
Private Const OPT_PLACE_MOVE As String = "셀과 함께 이동"
Private Const OPT_PLACE_FIXED As String = "고정"

Private Const MSG_PROTECTED As String = "시트가 보호되어 있어 사진을 넣거나 옮길 수 없습니다." & vbCrLf & _
                                        "검토 > 시트 보호 해제 후 다시 실행하세요."
Private Const MSG_BAD_AREA As String = "사진을 넣을 수 없는 칸입니다. (숨겨졌거나 너무 작은 칸)"

Private Type FitOptions
    MarginPt As Double
    KeepRatio As Boolean
    AlignCenter As Boolean
    WholeSelection As Boolean
    Placement As Long
End Type

Private mKeyInsert As String
Private mKeyPaste As String

'--- 매크로 1: 사진 파일 넣기 ------------------------------------
Public Sub ImageFit_InsertPicture()
    Dim ws As Worksheet, area As Range, o As FitOptions
    Dim p As String, why As String, msg As String
    Dim shp As Shape, st As CoAppState

    o = LoadOptions()
    If Not PickArea(ws, area, why, o) Then CoGuard_Notice why: Exit Sub
    If CoRange_IsProtected(ws, True) Then CoGuard_Notice MSG_PROTECTED: Exit Sub

    p = CoFile_PickFile("칸에 넣을 사진 선택", CO_FILTER_PICTURES)
    If Len(p) = 0 Then Exit Sub                         ' 취소

    st = CoGuard_Begin()
    On Error GoTo EH
    Set shp = CoShape_AddPicture(ws, p, why)
    If shp Is Nothing Then CoGuard_End st: CoGuard_Notice why: Exit Sub
    If Not FitShape(shp, area, o, False) Then
        CoShape_Delete shp
        CoGuard_End st
        CoGuard_Notice MSG_BAD_AREA
        Exit Sub
    End If
    CoGuard_End st
    Exit Sub
EH:
    msg = CoGuard_ErrText("사진을 넣는 중")
    CoShape_Delete shp                                  ' 반쯤 놓인 사진을 남기지 않는다
    CoGuard_End st
    CoGuard_Notice msg
End Sub

'--- 매크로 2: 붙여넣고 맞추기 -------------------------------------
Public Sub ImageFit_PasteAndFit()
    Dim ws As Worksheet, area As Range, o As FitOptions
    Dim why As String, msg As String, before As Long, pasteErr As Long
    Dim shp As Shape, st As CoAppState

    o = LoadOptions()
    If Not PickArea(ws, area, why, o) Then CoGuard_Notice why: Exit Sub   ' 붙이기 전에 칸부터 확인한다
    If CoRange_IsProtected(ws, True) Then CoGuard_Notice MSG_PROTECTED: Exit Sub

    st = CoGuard_Begin()
    On Error GoTo EH
    before = ws.Shapes.Count
    On Error Resume Next
    ws.Pictures.Paste                                   ' 사진이 없으면 오류 -> 아래에서 안내
    pasteErr = Err.Number
    Err.Clear
    On Error GoTo EH
    Set shp = CoShape_LastAdded(ws, before)
    If pasteErr <> 0 Or shp Is Nothing Then
        CoGuard_End st
        CoGuard_Notice "클립보드에 붙일 사진이 없습니다." & vbCrLf & "사진을 복사(Ctrl+C)한 뒤 다시 실행하세요."
        Exit Sub
    End If
    If Not FitShape(shp, area, o, False) Then
        CoShape_Delete shp
        CoGuard_End st
        CoGuard_Notice MSG_BAD_AREA
        Exit Sub
    End If
    CoGuard_End st
    Exit Sub
EH:
    msg = CoGuard_ErrText("붙여넣는 중")
    CoGuard_End st
    CoGuard_Notice msg
End Sub

'--- 매크로 3: 올려 둔 사진을 칸에 맞추기 -----------------------------
Public Sub ImageFit_FitSelected()
    Dim ws As Worksheet, o As FitOptions, sr As ShapeRange, area As Range
    Dim why As String, msg As String, i As Long, done As Long, skipped As Long
    Dim st As CoAppState

    If Not CoRange_ActiveWorksheet(ws, why) Then CoGuard_Notice why: Exit Sub
    If CoRange_IsProtected(ws, True) Then CoGuard_Notice MSG_PROTECTED: Exit Sub
    If Not CoShape_GetSelected(sr) Then CoGuard_Notice "맞출 사진을 먼저 선택하세요.": Exit Sub

    o = LoadOptions()
    st = CoGuard_Begin()
    On Error GoTo EH
    For i = 1 To sr.Count
        Set area = sr(i).TopLeftCell.MergeArea          ' 사진 왼쪽 위가 놓인 칸(병합이면 병합 전체)
        If CoShape_IsUpright(sr(i)) And CoRange_IsVisible(area) Then
            If FitShape(sr(i), area, o, True) Then done = done + 1 Else skipped = skipped + 1
        Else
            skipped = skipped + 1
        End If
    Next i
    CoGuard_End st
    If skipped > 0 Then
        CoGuard_Notice done & "개를 맞췄고 " & skipped & "개는 건너뛰었습니다." & vbCrLf & _
                       "(회전된 사진이거나 숨겨진 칸에 놓인 사진은 건너뜁니다)"
    End If
    Exit Sub
EH:
    msg = CoGuard_ErrText("맞추는 중")
    CoGuard_End st
    CoGuard_Notice msg
End Sub

'--- 단축키 (ThisWorkbook 코드에서 열 때·닫을 때 호출) ------------------
Public Sub ImageFit_RegisterShortcuts()
    On Error Resume Next                                ' 단축키 실패가 추가기능 로드를 막지 않게
    ImageFit_UnregisterShortcuts
    mKeyInsert = CoSetting_Text(ADDIN_ID, "keyInsert", "^+A")
    mKeyPaste = CoSetting_Text(ADDIN_ID, "keyPaste", "^+V")
    If Len(mKeyInsert) > 0 Then Application.OnKey mKeyInsert, "'" & ThisWorkbook.Name & "'!ImageFit_InsertPicture"
    If Len(mKeyPaste) > 0 Then Application.OnKey mKeyPaste, "'" & ThisWorkbook.Name & "'!ImageFit_PasteAndFit"
End Sub

Public Sub ImageFit_UnregisterShortcuts()
    On Error Resume Next
    If Len(mKeyInsert) > 0 Then Application.OnKey mKeyInsert   ' 인자 없이 호출하면 원래 동작으로 복구
    If Len(mKeyPaste) > 0 Then Application.OnKey mKeyPaste
    mKeyInsert = "": mKeyPaste = ""
End Sub

'--- 이 추가기능만의 도우미 ---------------------------------------
Private Function LoadOptions() As FitOptions
    Dim o As FitOptions
    o.MarginPt = CoSetting_MmAsPt(ADDIN_ID, "margin", 0.8, 0, 20)      ' JSON의 default·min·max와 같게
    o.KeepRatio = CoSetting_Bool(ADDIN_ID, "keepRatio", True)
    o.AlignCenter = (CoSetting_Choice(ADDIN_ID, "align", OPT_ALIGN_CENTER, OPT_ALIGN_CENTER, OPT_ALIGN_TOPLEFT) <> OPT_ALIGN_TOPLEFT)
    o.WholeSelection = (CoSetting_Choice(ADDIN_ID, "fitArea", OPT_AREA_ALL, OPT_AREA_ALL, OPT_AREA_FIRSTCELL) <> OPT_AREA_FIRSTCELL)
    Select Case CoSetting_Choice(ADDIN_ID, "placement", OPT_PLACE_ALL, OPT_PLACE_ALL, OPT_PLACE_MOVE, OPT_PLACE_FIXED)
        Case OPT_PLACE_MOVE: o.Placement = xlMove
        Case OPT_PLACE_FIXED: o.Placement = xlFreeFloating
        Case Else: o.Placement = xlMoveAndSize
    End Select
    LoadOptions = o
End Function

' 선택에서 사진을 맞출 칸을 정한다. 안 되면 False와 이유(why).
Private Function PickArea(ByRef ws As Worksheet, ByRef area As Range, ByRef why As String, ByRef o As FitOptions) As Boolean
    Dim rng As Range
    If Not CoRange_GetSelection(ws, rng, why) Then Exit Function      ' 떨어진 선택은 첫 덩어리, 행·열 전체는 거부
    If o.WholeSelection Then Set area = rng Else Set area = CoRange_FirstCellArea(rng)
    If Not CoRange_IsVisible(area) Then why = "선택한 칸이 숨겨져 있거나 크기가 0입니다.": Exit Function
    PickArea = True
End Function

' 사진을 칸에 맞춘다. 다시 맞출 때(restore)는 먼저 원본 비율을 되찾는다.
Private Function FitShape(ByVal shp As Shape, ByVal area As Range, ByRef o As FitOptions, ByVal restore As Boolean) As Boolean
    If restore And o.KeepRatio Then CoShape_RestoreOriginalSize shp
    FitShape = CoShape_FitToRange(shp, area, o.MarginPt, o.KeepRatio, o.AlignCenter, o.Placement)
End Function
