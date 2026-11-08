VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00A3F34D} SelStatsForm 
   Caption         =   "선택 요약"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3600
   StartUpPosition =   0  '수동
End
Attribute VB_Name = "SelStatsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================
' 기능: 선택 요약 창. 이름 글자와 큰 숫자 글자를 줄마다 만들어 보여 주는 화면만 맡는다(계산은 SelStatsCalc.bas).
' 사용법: Build(줄 이름 배열, 큰 글자 크기)로 한 번 만든 뒤 SetValue(줄 번호, 글자)로 값을 바꾼다.
'         컨트롤은 코드로 만들어서 .frx 파일이 필요 없다.
' 의존: SelStats.bas (창을 닫으면 SelStats_FormClosed를 부른다)
'==============================================================
Option Explicit

Private mValues() As MSForms.Label
Private mCount As Long

Public Sub Build(ByRef captions As Variant, ByVal bigPt As Double)
    Dim i As Long, nm As MSForms.Label, rowH As Double, y As Double, namePt As Double, w As Double
    mCount = UBound(captions) - LBound(captions) + 1
    ReDim mValues(1 To mCount)
    namePt = bigPt * 0.45: If namePt < 9 Then namePt = 9
    rowH = bigPt * 1.5
    w = bigPt * 9                                      ' 값 칸 너비(큰 글자 약 9자)
    y = 6
    For i = 1 To mCount
        Set nm = Me.Controls.Add("Forms.Label.1", "capt" & i, True)
        nm.Caption = captions(LBound(captions) + i - 1)
        nm.Font.Size = namePt
        nm.ForeColor = RGB(96, 96, 96)
        nm.Left = 10: nm.Top = y + (rowH - namePt * 1.4) / 2 + 2: nm.Width = namePt * 6: nm.Height = namePt * 1.5
        Set mValues(i) = Me.Controls.Add("Forms.Label.1", "val" & i, True)
        mValues(i).Font.Size = bigPt
        mValues(i).Font.Bold = True
        mValues(i).TextAlign = fmTextAlignRight
        mValues(i).Left = 10 + namePt * 6: mValues(i).Top = y: mValues(i).Width = w: mValues(i).Height = rowH
        y = y + rowH
    Next i
    Me.Width = 10 + namePt * 6 + w + 22
    Me.Height = y + 34
End Sub

Public Sub SetValue(ByVal i As Long, ByVal text As String)
    If i >= 1 And i <= mCount Then mValues(i).Caption = text
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    SelStats_FormClosed
End Sub
