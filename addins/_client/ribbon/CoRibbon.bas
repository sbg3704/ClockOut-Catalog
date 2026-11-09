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
    CoRibbon_Log "click " & macroName & " @ " & ThisWorkbook.Name
    Application.StatusBar = "칼퇴: " & macroName & " 실행 중..."
    Application.Run "'" & ThisWorkbook.Name & "'!" & macroName
    Application.StatusBar = "칼퇴: " & macroName & " 실행했습니다."
    Exit Sub
EH:
    Application.StatusBar = False
    CoRibbon_Log "error " & macroName & ": " & Err.Number & " " & Err.Description
    MsgBox "'" & macroName & "' 명령을 실행하지 못했습니다." & vbCrLf & Err.Description, vbExclamation, "칼퇴"
End Sub

Public Sub CoRibbon_Load(ribbon As IRibbonUI)
    ' 진단용: 이 파일의 리본이 올라올 때 한 줄 남긴다(엑셀이 이 파일의 콜백을 찾을 수 있다는 증거).
    CoRibbon_Log "load @ " & ThisWorkbook.Name
End Sub

Private Sub CoRibbon_Log(ByVal msg As String)
    ' 진단용: 단추가 눌렸는지 확인할 수 있게 임시 폴더에 한 줄씩 남긴다(실패해도 무시).
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open Environ$("TEMP") & "\clockout_ribbon.log" For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & msg
    Close #f
End Sub
