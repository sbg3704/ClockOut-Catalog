VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrossHairEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================
' 기능: 엑셀 이벤트(선택 바뀜·시트 바뀜·저장·닫기)를 받아 CrossHair 모듈에 넘긴다.
' 사용법: CrossHair_On이 만들어 App에 Application을 넣는다. 로직은 모두 CrossHair.bas에 있다.
' 의존: CrossHair.bas
'==============================================================
Option Explicit

Public WithEvents App As Application

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    CrossHair_Apply Sh
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object)
    CrossHair_Apply Sh
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    On Error Resume Next
    CrossHair_Apply Wb.ActiveSheet
End Sub

Private Sub App_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    CrossHair_Strip Wb
    Application.OnTime Now, "'" & ThisWorkbook.Name & "'!CrossHair_Reapply"    ' 저장이 끝난 뒤 다시 건다
End Sub

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    CrossHair_Strip Wb
    Application.OnTime Now, "'" & ThisWorkbook.Name & "'!CrossHair_Reapply"    ' 닫기를 취소했으면 다시 건다
End Sub
