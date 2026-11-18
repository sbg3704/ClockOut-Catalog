VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetNavEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================
' 기능: 엑셀 이벤트(시트·통합 문서 바뀜, 창 크기 바뀜)를 받아 SheetNav 모듈에 넘긴다.
' 사용법: SheetNav_On이 만들어 App에 Application을 넣는다. 로직은 모두 SheetNav.bas에 있다.
' 의존: SheetNav.bas
'==============================================================
Option Explicit

Public WithEvents App As Application

Private Sub App_SheetActivate(ByVal Sh As Object)
    SheetNav_Refresh
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    SheetNav_Refresh
End Sub

Private Sub App_WorkbookNewSheet(ByVal Wb As Workbook, ByVal Sh As Object)
    SheetNav_Refresh
End Sub

Private Sub App_WindowResize(ByVal Wb As Workbook, ByVal Wn As Window)
    SheetNav_Dock
End Sub
