VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SheetNavForm 
   Caption         =   "시트 목록"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2800
   StartUpPosition =   0  '수동
End
Attribute VB_Name = "SheetNavForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================
' 기능: 시트 사이드바 창. 검색 칸, 시트 목록, 숨김 표시·반대쪽·새로고침 단추를 보여 주는 화면만 맡는다.
' 사용법: Build(글자 크기, 숨김 시트 보기)로 한 번 만든 뒤 SetItems(...)로 목록을 바꾸고, SetBounds(...)로 자리를 잡는다.
'         컨트롤은 코드로 만들어서 .frx 파일이 필요 없다. build_addin.ps1은 이 파일에서 폼을 새로 만들고 코드를 넣는다.
' 의존: SheetNav.bas (이동·새로고침·쪽 바꾸기·닫힘은 SheetNav_* 를 부른다)
'==============================================================
Option Explicit

Private WithEvents mSearch As MSForms.TextBox
Private WithEvents mList As MSForms.ListBox
Private WithEvents mHid As MSForms.CheckBox
Private WithEvents mSwap As MSForms.CommandButton
Private WithEvents mRef As MSForms.CommandButton
Private mHead As MSForms.Label
Private mInfo As MSForms.Label

Private mWb As String
Private mNames() As String
Private mHidden() As Boolean
Private mCount As Long
Private mActive As String
Private mFilt() As String
Private mFiltN As Long
Private mBusy As Boolean

Public Sub Build(ByVal fontPt As Double, ByVal showHidden As Boolean)
    Me.StartUpPosition = 0
    Me.Caption = "시트 목록"
    Set mHead = Me.Controls.Add("Forms.Label.1", "lblWb", True)
    mHead.Font.Size = fontPt: mHead.Font.Bold = True: mHead.ForeColor = RGB(64, 64, 64)
    Set mSearch = Me.Controls.Add("Forms.TextBox.1", "txtFind", True)
    mSearch.Font.Size = fontPt: mSearch.ControlTipText = "시트 이름 검색 (Enter: 첫 결과로 이동, 아래 화살표: 목록으로)"
    Set mList = Me.Controls.Add("Forms.ListBox.1", "lstSheets", True)
    mList.Font.Size = fontPt: mList.MultiSelect = fmMultiSelectSingle
    Set mInfo = Me.Controls.Add("Forms.Label.1", "lblInfo", True)
    mInfo.Font.Size = fontPt - 1: mInfo.ForeColor = RGB(110, 110, 110)
    Set mHid = Me.Controls.Add("Forms.CheckBox.1", "chkHidden", True)
    mHid.Font.Size = fontPt - 1: mHid.Caption = "숨김 시트도 보기": mHid.Value = showHidden
    Set mSwap = Me.Controls.Add("Forms.CommandButton.1", "cmdSide", True)
    mSwap.Font.Size = fontPt - 1: mSwap.Caption = "반대쪽"
    Set mRef = Me.Controls.Add("Forms.CommandButton.1", "cmdRefresh", True)
    mRef.Font.Size = fontPt - 1: mRef.Caption = "새로고침"
    Layout
End Sub

' 자리와 크기를 잡고 안의 컨트롤을 다시 배치한다.
Public Sub SetBounds(ByVal l As Double, ByVal t As Double, ByVal w As Double, ByVal h As Double)
    Me.Left = l: Me.Top = t: Me.Width = w: Me.Height = h
    Layout
End Sub

Private Sub Layout()
    Dim w As Double, h As Double, lh As Double
    On Error Resume Next
    w = Me.InsideWidth: h = Me.InsideHeight
    mHead.Left = 6: mHead.Top = 4: mHead.Width = w - 12: mHead.Height = 14
    mSearch.Left = 6: mSearch.Top = 22: mSearch.Width = w - 12: mSearch.Height = 20
    lh = h - 46 - 66: If lh < 40 Then lh = 40
    mList.Left = 6: mList.Top = 46: mList.Width = w - 12: mList.Height = lh
    mInfo.Left = 6: mInfo.Top = 46 + lh + 4: mInfo.Width = w - 12: mInfo.Height = 14
    mHid.Left = 6: mHid.Top = 46 + lh + 20: mHid.Width = w - 12: mHid.Height = 16
    mSwap.Left = 6: mSwap.Top = 46 + lh + 40: mSwap.Width = (w - 18) / 2: mSwap.Height = 20
    mRef.Left = 12 + (w - 18) / 2: mRef.Top = mSwap.Top: mRef.Width = (w - 18) / 2: mRef.Height = 20
End Sub

' 목록을 새로 넣는다. names/hidden은 1부터 count까지, activeName은 지금 보고 있는 시트(목록에서 골라 둔다).
Public Sub SetItems(ByVal wbName As String, ByVal names As Variant, ByVal hidden As Variant, ByVal count As Long, ByVal activeName As String)
    Dim i As Long
    mWb = wbName: mCount = count: mActive = activeName
    If count > 0 Then
        ReDim mNames(1 To count): ReDim mHidden(1 To count)
        For i = 1 To count
            mNames(i) = names(i): mHidden(i) = hidden(i)
        Next i
    End If
    mHead.Caption = IIf(Len(wbName) = 0, "(열린 통합 문서 없음)", wbName)
    Fill
End Sub

' 검색어에 맞는 시트만 목록에 채우고, 보고 있는 시트를 골라 둔다.
Private Sub Fill()
    Dim i As Long, q As String, shown As String
    mBusy = True
    On Error Resume Next
    q = mSearch.Text
    mList.Clear
    mFiltN = 0
    For i = 1 To mCount
        If SheetNav_Match(mNames(i), q) Then
            mFiltN = mFiltN + 1
            ReDim Preserve mFilt(1 To mFiltN)
            mFilt(mFiltN) = mNames(i)
            shown = mNames(i)
            If mHidden(i) Then shown = shown & "  (숨김)"
            mList.AddItem shown
            If StrComp(mNames(i), mActive, vbTextCompare) = 0 Then mList.ListIndex = mFiltN - 1
        End If
    Next i
    If Len(q) > 0 Then
        mInfo.Caption = "검색 결과 " & mFiltN & "개 / 전체 " & mCount & "개"
    Else
        mInfo.Caption = "시트 " & mCount & "개"
    End If
    mBusy = False
End Sub

Private Sub NavTo(ByVal idx As Long)
    If idx >= 1 And idx <= mFiltN Then SheetNav_Go mWb, mFilt(idx)
End Sub

Private Sub mSearch_Change()
    Fill
End Sub

Private Sub mSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        NavTo 1
    ElseIf KeyCode = vbKeyDown Then
        KeyCode = 0
        If mList.ListCount > 0 Then mList.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = 0
        mSearch.Text = ""
    End If
End Sub

Private Sub mList_Click()
    If mBusy Then Exit Sub
    NavTo mList.ListIndex + 1
End Sub

Private Sub mHid_Click()
    SheetNav_SetShowHidden mHid.Value
End Sub

Private Sub mSwap_Click()
    SheetNav_SwapSide
End Sub

Private Sub mRef_Click()
    SheetNav_Refresh
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    SheetNav_FormClosed
End Sub
