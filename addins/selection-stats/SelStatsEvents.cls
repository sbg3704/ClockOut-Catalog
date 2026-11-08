VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelStatsEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================
' 기능: 엑셀 이벤트(선택 바뀜·값 바뀜·시트·통합 문서 바뀜)를 받아 SelStats 모듈에 넘긴다.
' 사용법: SelStats_On이 만들어 App에 Application을 넣는다. 로직은 모두 SelStats.bas에 있다.
' 의존: SelStats.bas
'==============================================================
Option Explicit

Public WithEvents App As Application

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    SelStats_Update
End Sub

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    SelStats_Update
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object)
    SelStats_Update
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    SelStats_Update
End Sub
