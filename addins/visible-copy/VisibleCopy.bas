Attribute VB_Name = "VisibleCopy"
'@clockout id=visible-copy name=보이는 칸만 복사·붙여넣기 host=excel
'@desc 필터·숨긴 줄/열이 있을 때 보이는 칸만 복사하고, 보이는 칸에만 붙여넣습니다(엑셀 기본은 숨긴 줄까지 덮어씀).
'==============================================================
' 기능: 엑셀은 필터·숨김이 걸린 곳에 붙여넣으면 숨은 줄에도 값을 덮어쓰고, 손으로 숨긴 줄·열은 복사에도 같이 딸려 온다.
'       (1) 보이는 칸만 복사해 두기  (2) 복사해 둔 것을 목적지의 보이는 줄·열에만 차례로 붙여넣기.
' 사용법: 우클릭 메뉴(또는 리본 버튼)에서 "보이는 칸만 복사", 붙일 첫 칸에서 "보이는 칸에만 붙여넣기". 단축키는 기본 없음(설정에서 지정).
'         셀 우클릭 > "칼퇴 보이는 칸만 복사·붙여넣기"에도 같은 항목이 있다. 단축키는 설정에서 바꾼다(비우면 없음).
' 설정: vcPasteMode("전부"=서식·수식 포함 / "값만"). 단축키는 설정 key.copy, key.paste(Application.OnKey 형식, ^=Ctrl +=Shift %=Alt).
' 규칙: 붙여넣기 전에 줄·칸 수를 알려 주고 확인을 받는다. 목적지에서 숨은 줄·열은 건드리지 않는다. 실행 후 Ctrl+Z 불가.
'       원본의 보이는 줄·열 수만큼 목적지에 보이는 줄·열이 있어야 한다(모자라면 하지 않는다). 최대 100000칸.
' 의존: 공용 라이브러리 clockout-lib 0.1.001 이상(CoRange_*, CoSetting_*, CoGuard_*). CoLibClient.bas가 연결한다.
' 시험: Test_VisibleCopy (보이는 줄 번호 모으기 규칙)
'==============================================================
Option Explicit
Option Private Module

Public Const VC_ADDIN_ID As String = "visible-copy"
Private Const ADDIN_NAME As String = "보이는 칸만 복사·붙여넣기"
Private Const MIN_LIB As String = "0.1.001"
Private Const MENU_TAG As String = "ClockOutVisibleCopy"
Private Const MAX_CELLS As Long = 100000

Private mSrc As Range          ' 복사해 둔 원본 범위
Private mKeyCopy As String
Private mKeyPaste As String

Public Sub VisibleCopy_Startup()
    Dim api As Object, pop As Object
    On Error Resume Next
    VisibleCopy_Shutdown
    Set pop = Application.CommandBars("Cell").Controls.Add(Type:=10, Temporary:=True)
    pop.Caption = "칼퇴 보이는 칸만 복사·붙여넣기"
    pop.Tag = MENU_TAG
    With pop.Controls.Add(Type:=1, Temporary:=True)
        .Caption = "보이는 칸만 복사"
        .OnAction = "'" & ThisWorkbook.Name & "'!VisibleCopy_Copy"
    End With
    With pop.Controls.Add(Type:=1, Temporary:=True)
        .Caption = "보이는 칸에만 붙여넣기"
        .OnAction = "'" & ThisWorkbook.Name & "'!VisibleCopy_Paste"
    End With
    Set api = CoLibClient_Api(MIN_LIB, ADDIN_NAME, True)
    If api Is Nothing Then Exit Sub
    mKeyCopy = api.CoSetting_Text(VC_ADDIN_ID, "key.copy", "")
    mKeyPaste = api.CoSetting_Text(VC_ADDIN_ID, "key.paste", "")
    If Len(mKeyCopy) > 0 Then Application.OnKey mKeyCopy, "'" & ThisWorkbook.Name & "'!VisibleCopy_Copy"
    If Len(mKeyPaste) > 0 Then Application.OnKey mKeyPaste, "'" & ThisWorkbook.Name & "'!VisibleCopy_Paste"
End Sub

Public Sub VisibleCopy_Shutdown()
    Dim i As Long, ctl As Object
    On Error Resume Next
    For i = Application.CommandBars("Cell").Controls.Count To 1 Step -1
        Set ctl = Application.CommandBars("Cell").Controls(i)
        If ctl.Tag = MENU_TAG Then ctl.Delete
    Next i
    If Len(mKeyCopy) > 0 Then Application.OnKey mKeyCopy
    If Len(mKeyPaste) > 0 Then Application.OnKey mKeyPaste
    mKeyCopy = "": mKeyPaste = ""
    Set mSrc = Nothing
End Sub

'--- (1) 보이는 칸만 복사 ----------------------------------------
Public Sub VisibleCopy_Copy()
    Dim api As Object, ws As Worksheet, sel As Range, why As String, rows1 As Collection, cols1 As Collection
    Set api = CoLibClient_Api(MIN_LIB, ADDIN_NAME)
    If api Is Nothing Then Exit Sub
    If Not api.CoRange_GetSelection(ws, sel, why, False, False) Then api.CoGuard_Notice why: Exit Sub
    Set rows1 = VisibleIndexes(sel, True)
    Set cols1 = VisibleIndexes(sel, False)
    If rows1.Count = 0 Or cols1.Count = 0 Then api.CoGuard_Notice "보이는 칸이 없습니다.": Exit Sub
    If CLng(rows1.Count) * CLng(cols1.Count) > MAX_CELLS Then api.CoGuard_Notice "칸이 너무 많습니다(최대 " & MAX_CELLS & "칸).": Exit Sub
    Set mSrc = sel
    sel.SpecialCells(xlCellTypeVisible).Copy                  ' 엑셀 붙여넣기(다른 프로그램)용으로도 보이는 칸만 담는다
    Application.StatusBar = "보이는 칸만 복사: " & rows1.Count & "줄 x " & cols1.Count & "칸 (숨은 " & (sel.Rows.Count - rows1.Count) & "줄·" & (sel.Columns.Count - cols1.Count) & "열 제외)"
End Sub

' 범위 안에서 보이는 줄(또는 열)의 위치 번호(1부터)를 모은다. 필터로 숨겨진 줄도 Hidden으로 잡힌다.
Public Function VisibleIndexes(ByVal rng As Range, ByVal byRow As Boolean) As Collection
    Dim out As New Collection, i As Long, n As Long
    If byRow Then n = rng.Rows.Count Else n = rng.Columns.Count
    For i = 1 To n
        If byRow Then
            If Not rng.Rows(i).EntireRow.Hidden Then out.Add i
        Else
            If Not rng.Columns(i).EntireColumn.Hidden Then out.Add i
        End If
    Next i
    Set VisibleIndexes = out
End Function

'--- (2) 보이는 칸에만 붙여넣기 ------------------------------------
Public Sub VisibleCopy_Paste()
    Dim api As Object, ws As Worksheet, sel As Range, why As String, sr As Collection, sc As Collection
    Dim dr As Collection, dc As Collection, valuesOnly As Boolean, i As Long, j As Long, tok As Long, msg As String
    Dim lastR As Long, lastC As Long, c As Long, probe As Range, src As Range, dst As Range
    Set api = CoLibClient_Api(MIN_LIB, ADDIN_NAME)
    If api Is Nothing Then Exit Sub
    If mSrc Is Nothing Then api.CoGuard_Notice "먼저 '보이는 칸만 복사'를 해 두세요.": Exit Sub
    On Error Resume Next
    Set probe = mSrc.Cells(1, 1)
    If Err.Number <> 0 Then Err.Clear: Set mSrc = Nothing: api.CoGuard_Notice "복사해 둔 원본이 닫혔습니다. 다시 복사하세요.": Exit Sub
    On Error GoTo 0
    If Not api.CoRange_GetSelection(ws, sel, why, False, False) Then api.CoGuard_Notice why: Exit Sub
    Set sr = VisibleIndexes(mSrc, True)
    Set sc = VisibleIndexes(mSrc, False)
    ' 목적지의 보이는 줄·열을 시작 칸부터 필요한 만큼 모은다
    Set dr = CollectVisible(ws, sel.Row, sr.Count, True)
    Set dc = CollectVisible(ws, sel.Column, sc.Count, False)
    If dr.Count < sr.Count Or dc.Count < sc.Count Then
        api.CoGuard_Notice "목적지에 보이는 칸이 모자랍니다(필요 " & sr.Count & "줄 x " & sc.Count & "칸, 있는 칸 " & dr.Count & "줄 x " & dc.Count & "칸)."
        Exit Sub
    End If
    lastR = dr(dr.Count): lastC = dc(dc.Count)
    Set dst = ws.Range(ws.Cells(dr(1), dc(1)), ws.Cells(lastR, lastC))
    If Not api.CoRange_CanEdit(dst) Then api.CoGuard_Notice "시트가 보호되어 있고 잠긴 칸이 있어 붙여넣을 수 없습니다.": Exit Sub
    valuesOnly = (api.CoSetting_Choice(VC_ADDIN_ID, "vcPasteMode", "전부", "전부", "값만") = "값만")
    If Not api.CoGuard_Ask(sr.Count & "줄 x " & sc.Count & "칸을 '" & ws.Cells(dr(1), dc(1)).Address(False, False) & "'부터 보이는 칸에만 붙여넣습니다." & vbCrLf & _
        "(" & IIf(valuesOnly, "값만", "서식·수식 포함") & ", 목적지의 숨은 줄·열 " & ((lastR - dr(1) + 1 - dr.Count) + (lastC - dc(1) + 1 - dc.Count)) & "개는 건드리지 않음)" & vbCrLf & vbCrLf & _
        "실행 후에는 Ctrl+Z로 되돌릴 수 없습니다. 진행할까요?") Then Exit Sub
    tok = api.CoGuard_Begin(True, True)
    On Error GoTo EH
    For i = 1 To sr.Count
        For j = 1 To sc.Count
            Set src = mSrc.Cells(sr(i), sc(j))
            Set dst = ws.Cells(dr(i), dc(j))
            If valuesOnly Then dst.Value2 = src.Value2 Else src.Copy Destination:=dst
        Next j
    Next i
    Application.CutCopyMode = False
    api.CoGuard_End tok
    api.CoGuard_Notice sr.Count & "줄 x " & sc.Count & "칸을 붙여넣었습니다."
    Exit Sub
EH:
    msg = "붙여넣는 중 문제가 생겼습니다: " & Err.Description
    Err.Clear
    Application.CutCopyMode = False
    api.CoGuard_End tok
    api.CoGuard_Notice msg
End Sub

' 시작 줄(열)부터 보이는 줄(열) 번호를 need개까지 모은다(시트 끝에서 멈춘다).
Private Function CollectVisible(ByVal ws As Worksheet, ByVal startAt As Long, ByVal need As Long, ByVal byRow As Boolean) As Collection
    Dim out As New Collection, k As Long, lim As Long
    If byRow Then lim = ws.Rows.Count Else lim = ws.Columns.Count
    k = startAt
    Do While k <= lim And out.Count < need
        If byRow Then
            If Not ws.Rows(k).Hidden Then out.Add k
        Else
            If Not ws.Columns(k).Hidden Then out.Add k
        End If
        k = k + 1
    Loop
    Set CollectVisible = out
End Function

'--- 시험 ----------------------------------------------------------
Public Sub Test_VisibleCopy()
    Dim wb As Workbook, ws As Worksheet, r As Collection, f As Long
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Rows(3).Hidden = True
    ws.Rows(5).Hidden = True
    Set r = VisibleIndexes(ws.Range("A2:A6"), True)
    If r.Count <> 3 Then f = f + 1
    If r(1) <> 1 Or r(2) <> 3 Or r(3) <> 5 Then f = f + 1
    wb.Close SaveChanges:=False
    If f = 0 Then
        MsgBox "VisibleCopy 시험: 모두 통과", vbInformation, "칼퇴"
    Else
        MsgBox "VisibleCopy 시험: 실패 " & f & "건", vbExclamation, "칼퇴"
    End If
End Sub
