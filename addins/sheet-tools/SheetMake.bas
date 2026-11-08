Attribute VB_Name = "SheetMake"
'==============================================================
' 기능: 선택한 칸의 글자(한 열이든 여러 칸이든)를 시트 이름으로 새 시트를 만든다.
' 사용법: 이름이 적힌 칸을 고르고 시트 탭 우클릭 > 칼퇴 시트 도구 > 목록으로 시트 만들기.
' 설정: makeBase("빈 시트" 또는 "현재 시트 복사"), makePosition("맨 뒤" 또는 "현재 시트 뒤").
' 규칙: 이미 있는 이름·빈 칸·중복은 건너뛴다. 금지 문자(: \ / ? * [ ])는 _로, 31자로 자른다.
' 의존: SheetTools.bas, 공용 라이브러리(CoText_SafeSheetName, CoRange_GetSelection, CoGuard_*)
'==============================================================
Option Explicit
Option Private Module

Public Sub SheetMake_FromList()
    Dim api As Object, ws As Worksheet, sel As Range, why As String, wb As Workbook
    Dim names As Collection, v As Variant, cel As Range, s As String, nm As String
    Dim copyIt As Boolean, afterCur As Boolean, tok As Long, made As Long, skipped As Long, msg As String
    Dim seen As Object, ar As Range, i As Long, j As Long, newWs As Worksheet
    Set api = ST_Api()
    If api Is Nothing Then Exit Sub
    If Not api.CoRange_GetSelection(ws, sel, why, True, False) Then api.CoGuard_Notice why: Exit Sub
    Set wb = ws.Parent
    If wb.ProtectStructure Then api.CoGuard_Notice "통합 문서 구조가 보호되어 있어 시트를 만들 수 없습니다.": Exit Sub
    copyIt = (api.CoSetting_Choice(ST_ADDIN_ID, "makeBase", "빈 시트", "빈 시트", "현재 시트 복사") = "현재 시트 복사")
    afterCur = (api.CoSetting_Choice(ST_ADDIN_ID, "makePosition", "맨 뒤", "맨 뒤", "현재 시트 뒤") = "현재 시트 뒤")
    Set names = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1
    For Each ar In sel.Areas
        v = api.CoRange_Values2D(ar)
        For i = 1 To UBound(v, 1)
            For j = 1 To UBound(v, 2)
                If Not IsError(v(i, j)) Then
                    s = Trim$(CStr(v(i, j)))
                    If Len(s) > 0 Then
                        nm = api.CoText_SafeSheetName(s)
                        If seen.Exists(nm) Or ST_SheetExists(wb, nm) Then
                            skipped = skipped + 1
                        Else
                            seen.Add nm, 1
                            names.Add nm
                        End If
                    End If
                End If
            Next j
        Next i
    Next ar
    If names.Count = 0 Then api.CoGuard_Notice "새로 만들 이름이 없습니다." & IIf(skipped > 0, vbCrLf & "(" & skipped & "개는 이미 있거나 중복이라 건너뜀)", ""): Exit Sub
    If names.Count > 100 Then
        If Not api.CoGuard_Ask("시트 " & names.Count & "개를 만듭니다. 계속할까요?") Then Exit Sub
    End If
    tok = api.CoGuard_Begin(True, True)
    On Error GoTo EH
    For i = 1 To names.Count
        If afterCur Then
            If copyIt Then ws.Copy After:=wb.Sheets(IIf(made = 0, ws.Index, ws.Index + made)) Else wb.Worksheets.Add After:=wb.Sheets(ws.Index + made)
        Else
            If copyIt Then ws.Copy After:=wb.Sheets(wb.Sheets.Count) Else wb.Worksheets.Add After:=wb.Sheets(wb.Sheets.Count)
        End If
        Set newWs = ActiveSheet
        newWs.Name = names(i)
        made = made + 1
    Next i
    api.CoGuard_End tok
    ws.Activate
    api.CoGuard_Notice made & "개 시트를 만들었습니다." & IIf(skipped > 0, vbCrLf & "(" & skipped & "개는 이미 있거나 중복이라 건너뜀)", "")
    Exit Sub
EH:
    msg = "시트를 만드는 중 문제가 생겼습니다(" & made & "개 만든 뒤): " & Err.Description
    Err.Clear
    api.CoGuard_End tok
    api.CoGuard_Notice msg
End Sub
