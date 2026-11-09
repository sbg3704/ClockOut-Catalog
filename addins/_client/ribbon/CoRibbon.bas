Attribute VB_Name = "CoRibbon"
'==============================================================
' 기능: 리본 "칼퇴" 탭의 버튼을 눌렀을 때 그 명령의 매크로를 실행한다.
'       선언(.clockout.json)에 commands가 있는 추가기능에만 build_addin.ps1이 자동으로 넣는다.
' 동작: 리본 XML의 각 버튼 tag에 매크로 이름이 들어 있고, 눌리면 이 추가기능 안의 그 매크로를 부른다.
' 규칙: 이 모듈은 직접 고치지 않는다(모든 추가기능에 똑같이 들어간다). 매크로는 공개(Public) Sub여야 한다.
' 의존: Office 개체 라이브러리(IRibbonControl) - 엑셀 VBA 프로젝트가 기본으로 참조한다.
'==============================================================
Option Explicit

Public Sub CoRibbon_Click(control As IRibbonControl)
    Dim macroName As String
    On Error GoTo EH
    macroName = control.Tag
    Application.Run "'" & ThisWorkbook.Name & "'!" & macroName
    Exit Sub
EH:
    MsgBox "'" & macroName & "' 명령을 실행하지 못했습니다." & vbCrLf & Err.Description, vbExclamation, "칼퇴"
End Sub
