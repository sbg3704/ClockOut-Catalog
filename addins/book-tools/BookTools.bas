Attribute VB_Name = "BookTools"
'@clockout id=book-tools name=통합 문서 도구 host=excel
'@desc 시트 탭 우클릭 "칼퇴 통합 문서 도구" 메뉴. 숨은 시트 보이기/선택 시트만 남기기, 공유 전 정리, 원클릭 백업, 선택 시트를 PDF 하나로, 입력 칸 비우기, 시트 이름순 정렬, 시트 이름 일괄 바꾸기.
'==============================================================
' 기능: 시트 탭 우클릭 메뉴("칼퇴 통합 문서 도구")를 만들고 도구들이 같이 쓰는 도우미를 둔다.
' 메뉴(각 도구는 다른 모듈에 있다):
'   숨은 시트 모두 보이기      BookTools_ShowAll        (BookSheets.bas)
'   선택한 시트만 남기고 숨기기  BookTools_HideOthers     (BookSheets.bas)
'   시트 탭 이름순 정렬        BookTools_SortTabs       (BookSheets.bas)
'   시트 이름 일괄 바꾸기      BookTools_RenameSheets   (BookSheets.bas)
'   공유 전 정리(A1·100%)      BookTools_PrepShare      (BookFiles.bas)
'   원클릭 백업(날짜시각 복사본) BookTools_Backup        (BookFiles.bas)
'   선택 시트를 PDF 하나로     BookTools_ExportPdf      (BookFiles.bas)
'   입력 칸 비우기(수식 유지)   BookTools_ClearInputs    (BookFiles.bas)
' 규칙: 지우거나 바꾸는 도구는 확인을 받는다. 설정은 칼퇴 설정창(book-tools.clockout.json).
' 의존: 공용 라이브러리 clockout-lib 0.1.001 이상. CoLibClient.bas가 연결한다.
'==============================================================
Option Explicit
Option Private Module

Public Const BT_ADDIN_ID As String = "book-tools"
Private Const ADDIN_NAME As String = "통합 문서 도구"
Private Const MIN_LIB As String = "0.1.001"
Private Const MENU_TAG As String = "ClockOutBookTools"

Public Sub BookTools_Startup()
    On Error Resume Next
    BookTools_Shutdown
    Dim pop As Object, items As Variant, i As Long
    items = Array("숨은 시트 모두 보이기|BookTools_ShowAll", _
                  "선택한 시트만 남기고 숨기기|BookTools_HideOthers", _
                  "시트 탭 이름순 정렬|BookTools_SortTabs", _
                  "시트 이름 일괄 바꾸기|BookTools_RenameSheets", _
                  "공유 전 정리(A1·배율 100%)|BookTools_PrepShare", _
                  "원클릭 백업(날짜시각 복사본)|BookTools_Backup", _
                  "선택한 시트를 PDF 하나로|BookTools_ExportPdf", _
                  "입력 칸 비우기(수식은 유지)|BookTools_ClearInputs")
    Set pop = Application.CommandBars("Ply").Controls.Add(Type:=10, Temporary:=True)
    pop.Caption = "칼퇴 통합 문서 도구"
    pop.Tag = MENU_TAG
    For i = LBound(items) To UBound(items)
        With pop.Controls.Add(Type:=1, Temporary:=True)
            .Caption = Split(items(i), "|")(0)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & Split(items(i), "|")(1)
        End With
    Next i
End Sub

Public Sub BookTools_Shutdown()
    Dim i As Long, ctl As Object
    On Error Resume Next
    For i = Application.CommandBars("Ply").Controls.Count To 1 Step -1
        Set ctl = Application.CommandBars("Ply").Controls(i)
        If ctl.Tag = MENU_TAG Then ctl.Delete
    Next i
End Sub

Public Function BT_Api() As Object
    Set BT_Api = CoLibClient_Api(MIN_LIB, ADDIN_NAME)
End Function

Public Function BT_Confirm(ByVal api As Object, ByVal what As String) As Boolean
    BT_Confirm = api.CoGuard_Ask(what & vbCrLf & vbCrLf & "실행 후에는 Ctrl+Z로 되돌릴 수 없습니다. 진행할까요?")
End Function

' 지금 선택된 시트들(Ctrl로 여러 장). 하나만 선택됐으면 그 시트 하나. 차트 시트도 포함.
Public Function BT_SelectedSheets() As Collection
    Dim out As New Collection, sh As Object
    On Error Resume Next
    For Each sh In ActiveWindow.SelectedSheets
        out.Add sh
    Next sh
    If out.Count = 0 Then out.Add ActiveSheet
    Set BT_SelectedSheets = out
End Function
