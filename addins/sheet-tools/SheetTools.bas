Attribute VB_Name = "SheetTools"
'@clockout id=sheet-tools name=시트 도구 host=excel
'@desc 시트 탭을 우클릭하면 나오는 "칼퇴 시트 도구" 메뉴. 시트 목차 만들기, 목록으로 시트 만들기, 시트별 파일로 내보내기, 폴더 파일 목록 만들기.
'==============================================================
' 기능: 시트 탭 우클릭 메뉴("칼퇴 시트 도구")를 만들고, 각 도구가 같이 쓰는 도우미를 둔다.
' 메뉴(각 도구는 다른 모듈에 있다):
'   시트 목차 만들기        SheetIndex_Make        (SheetIndex.bas)
'   목록으로 시트 만들기    SheetMake_FromList     (SheetMake.bas)
'   시트별 파일로 내보내기  SheetExport_Each       (SheetExport.bas)
'   폴더 파일 목록 만들기   FolderList_Make        (FolderList.bas)
' 규칙: 새 시트를 만드는 도구는 기존 시트를 지우지 않는다. 설정은 칼퇴 설정창(sheet-tools.clockout.json).
' 의존: 공용 라이브러리 clockout-lib 0.1.001 이상. CoLibClient.bas가 연결한다.
'==============================================================
Option Explicit
Option Private Module

Public Const ST_ADDIN_ID As String = "sheet-tools"
Private Const ADDIN_NAME As String = "시트 도구"
Private Const MIN_LIB As String = "0.1.001"
Private Const MENU_TAG As String = "ClockOutSheetTools"

Public Sub SheetTools_Startup()
    On Error Resume Next
    SheetTools_Shutdown
    Dim pop As Object, items As Variant, i As Long
    items = Array("시트 목차 만들기|SheetIndex_Make", _
                  "목록으로 시트 만들기|SheetMake_FromList", _
                  "시트별 파일로 내보내기|SheetExport_Each", _
                  "폴더 파일 목록 만들기|FolderList_Make")
    Set pop = Application.CommandBars("Ply").Controls.Add(Type:=10, Temporary:=True)
    pop.Caption = "칼퇴 시트 도구"
    pop.Tag = MENU_TAG
    For i = LBound(items) To UBound(items)
        With pop.Controls.Add(Type:=1, Temporary:=True)
            .Caption = Split(items(i), "|")(0)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & Split(items(i), "|")(1)
        End With
    Next i
End Sub

Public Sub SheetTools_Shutdown()
    Dim i As Long, ctl As Object
    On Error Resume Next
    For i = Application.CommandBars("Ply").Controls.Count To 1 Step -1
        Set ctl = Application.CommandBars("Ply").Controls(i)
        If ctl.Tag = MENU_TAG Then ctl.Delete
    Next i
End Sub

' 공용 라이브러리 개체(없거나 판이 낮으면 안내 후 Nothing).
Public Function ST_Api() As Object
    Set ST_Api = CoLibClient_Api(MIN_LIB, ADDIN_NAME)
End Function

' 통합 문서에 그 이름의 시트가 있으면 True.
Public Function ST_SheetExists(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim sh As Object
    For Each sh In wb.Sheets
        If StrComp(sh.Name, nm, vbTextCompare) = 0 Then ST_SheetExists = True: Exit Function
    Next sh
End Function

' 겹치지 않는 시트 이름(겹치면 _2, _3 ...을 붙이고 31자를 지킨다).
Public Function ST_UniqueSheetName(ByVal api As Object, ByVal wb As Workbook, ByVal wanted As String) As String
    Dim base As String, cand As String, n As Long, suffix As String
    base = api.CoText_SafeSheetName(wanted)
    cand = base
    Do While ST_SheetExists(wb, cand)
        n = n + 1
        suffix = "_" & (n + 1)
        cand = Left$(base, 31 - Len(suffix)) & suffix
    Loop
    ST_UniqueSheetName = cand
End Function

' 시트 이름을 수식 안에 쓸 수 있게 따옴표 처리('이름'!A1).
Public Function ST_Quote(ByVal nm As String) As String
    ST_Quote = "'" & Replace(nm, "'", "''") & "'"
End Function
