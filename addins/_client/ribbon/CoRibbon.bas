Attribute VB_Name = "CoRibbon"
'==============================================================
' 기능: 리본 "칼퇴" 탭의 버튼을 눌렀을 때 그 명령의 매크로를 실행한다.
'       선언(.clockout.json)에 commands가 있는 추가기능에만 build_addin.ps1이 자동으로 넣는다.
' 동작: 리본 XML의 각 버튼 tag에 매크로 이름이 들어 있고, 눌리면 이 추가기능 안의 그 매크로를 부른다.
'       단추는 반드시 일반 id여야 한다(idQ면 onAction이 불리지 않음, tools/build_addin.md).
' 규칙: 이 모듈은 직접 고치지 않는다(모든 추가기능에 똑같이 들어간다). 매크로는 공개(Public) Sub여야 한다.
' 의존: Office 개체 라이브러리(IRibbonControl) - 엑셀 VBA 프로젝트가 기본으로 참조한다.
'==============================================================
Option Explicit

Public Sub CoRibbon_Click(control As IRibbonControl)
    Dim macroName As String, running As String
    On Error GoTo EH
    macroName = control.Tag
    running = "칼퇴: " & macroName & " 실행 중..."
    Application.StatusBar = running
    Application.Run "'" & ThisWorkbook.Name & "'!" & macroName
    ClearIfMine running                                 ' 매크로가 결과를 띄웠으면 그대로 둔다
    Exit Sub
EH:
    ClearIfMine running
    MsgBox "'" & macroName & "' 명령을 실행하지 못했습니다." & vbCrLf & Err.Description, vbExclamation, "칼퇴"
End Sub

' 상태 표시줄이 아직 이 모듈이 띄운 글이면 엑셀 기본 표시로 돌린다.
Private Sub ClearIfMine(ByVal running As String)
    On Error Resume Next
    If VarType(Application.StatusBar) = vbString Then
        If Application.StatusBar = running Then Application.StatusBar = False
    End If
End Sub
