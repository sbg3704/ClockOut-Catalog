Attribute VB_Name = "BookClean"
'@clockout id=book-clean name=문서 정리 도구 host=excel
'@desc 시트 탭 우클릭 "칼퇴 문서 정리" 메뉴. 마지막 셀(Ctrl+End) 바로잡기, 외부 링크·깨진 이름 정리, 인쇄 설정 일괄 복사.
'==============================================================
' 기능: 시트 탭 우클릭 메뉴("칼퇴 문서 정리")를 만들고 도구들이 같이 쓰는 도우미를 둔다.
' 메뉴(각 도구는 다른 모듈에 있다):
'   마지막 셀 바로잡기        BookClean_FixLastCell   (BookCleanLast.bas)
'   외부 링크·깨진 이름 정리  BookClean_Links         (BookCleanLinks.bas)
'   인쇄 설정 일괄 복사       BookClean_CopyPage      (BookCleanPage.bas)
' 규칙: 지우거나 바꾸기 전에 항목과 개수를 알리고 확인을 받는다. 실행 후 Ctrl+Z로 되돌릴 수 없으니 먼저 저장해 두라고 안내한다.
' 의존: 공용 라이브러리 clockout-lib 0.1.001 이상. CoLibClient.bas가 연결한다.
'==============================================================
Option Explicit
Option Private Module

Public Const BC_ADDIN_ID As String = "book-clean"
Private Const ADDIN_NAME As String = "문서 정리 도구"
Private Const MIN_LIB As String = "0.1.001"
Private Const MENU_TAG As String = "ClockOutBookClean"

Public Sub BookClean_Startup()
    On Error Resume Next
    BookClean_Shutdown
    Dim pop As Object, items As Variant, i As Long
    items = Array("마지막 셀 바로잡기|BookClean_FixLastCell", _
                  "외부 링크·깨진 이름 정리|BookClean_Links", _
                  "인쇄 설정 일괄 복사|BookClean_CopyPage")
    Set pop = Application.CommandBars("Ply").Controls.Add(Type:=10, Temporary:=True)
    pop.Caption = "칼퇴 문서 정리"
    pop.Tag = MENU_TAG
    For i = LBound(items) To UBound(items)
        With pop.Controls.Add(Type:=1, Temporary:=True)
            .Caption = Split(items(i), "|")(0)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & Split(items(i), "|")(1)
        End With
    Next i
End Sub

Public Sub BookClean_Shutdown()
    Dim i As Long, ctl As Object
    On Error Resume Next
    For i = Application.CommandBars("Ply").Controls.Count To 1 Step -1
        Set ctl = Application.CommandBars("Ply").Controls(i)
        If ctl.Tag = MENU_TAG Then ctl.Delete
    Next i
End Sub

Public Function BC_Api() As Object
    Set BC_Api = CoLibClient_Api(MIN_LIB, ADDIN_NAME)
End Function

' 값을 바꾸기 전 확인(되돌릴 수 없음을 함께 알림).
Public Function BC_Confirm(ByVal api As Object, ByVal what As String) As Boolean
    BC_Confirm = api.CoGuard_Ask(what & vbCrLf & vbCrLf & "실행 후에는 Ctrl+Z로 되돌릴 수 없습니다. 먼저 저장해 두셨나요? 진행할까요?")
End Function
