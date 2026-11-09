Attribute VB_Name = "RangeSplit"
'==============================================================
' 기능: 표를 한 열의 값(공종·업체·동 등)별로 나눠, 그룹마다 새 시트(제목 줄 + 해당 줄)를 만들고 소계 시트도 만든다.
' 사용법: 표 안의 칸 하나를 누르거나 표를 고른 뒤, 나눌 기준 열의 칸을 활성 칸으로 두고 우클릭 > 칼퇴 범위 도구 > 그룹별 시트로 쪼개기.
'         첫 줄은 제목으로 본다. 기준 열은 활성 칸이 있는 열(선택 밖이면 첫 열).
' 설정: splitSummary(소계 시트 만들기, 기본 켬 - 그룹별 줄 수와 숫자 열 합계), splitMaxGroups(최대 그룹 수, 기본 50).
' 규칙: 원본 시트는 바뀌지 않는다. 그룹 값이 비어 있으면 "(빈 값)" 시트. 값은 서식째 복사한다. 시트 이름 금지 문자는 _로.
' 의존: RangeTools.bas, 공용 라이브러리(CoText_SafeSheetName, CoRange_Values2D, CoSetting_*, CoGuard_*)
'==============================================================
Option Explicit
Option Private Module

Public Sub RangeSplit_ByGroup()
    Dim api As Object, ws As Worksheet, rng As Range, wb As Workbook, keyCol As Long, v As Variant, i As Long, j As Long
    Dim groups As Object, order As Collection, k As String, maxG As Long, withSum As Boolean, tok As Long, msg As String
    Dim hdr As Range, rows1 As Range, grp As Variant, out As Worksheet, made As Long, nm As String, n As Long
    Dim sumWs As Worksheet, sr As Long, c As Long, colSum As Double, hasNum As Boolean, firstOut As Worksheet
    Set api = RT_Api()
    If api Is Nothing Then Exit Sub
    Set rng = RT_Area(api, ws)
    If rng Is Nothing Then Exit Sub
    If rng.Areas.Count > 1 Or rng.Rows.Count < 2 Then api.CoGuard_Notice "제목 줄과 데이터 줄이 있는 표를 한 덩어리로 고르세요.": Exit Sub
    Set wb = ws.Parent
    If wb.ProtectStructure Then api.CoGuard_Notice "통합 문서 구조가 보호되어 있어 시트를 만들 수 없습니다.": Exit Sub
    maxG = CLng(api.CoSetting_Number(RT_ADDIN_ID, "splitMaxGroups", 50, 2, 200))
    withSum = api.CoSetting_Bool(RT_ADDIN_ID, "splitSummary", True)
    keyCol = 1
    If Not Intersect(ActiveCell, rng) Is Nothing Then keyCol = ActiveCell.Column - rng.Column + 1
    v = api.CoRange_Values2D(rng.Columns(keyCol))
    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = 1
    Set order = New Collection
    For i = 2 To UBound(v, 1)
        If IsError(v(i, 1)) Then k = "(오류)" Else k = Trim$(CStr(v(i, 1)))
        If Len(k) = 0 Then k = "(빈 값)"
        If Not groups.Exists(k) Then groups.Add k, New Collection: order.Add k
        groups(k).Add i
    Next i
    If order.Count > maxG Then api.CoGuard_Notice "그룹이 " & order.Count & "개로 너무 많습니다(최대 " & maxG & "). 기준 열을 확인하거나 설정에서 늘리세요.": Exit Sub
    If order.Count < 2 Then api.CoGuard_Notice "나눌 그룹이 하나뿐입니다(기준 열: " & rng.Cells(1, keyCol).Value2 & ").": Exit Sub
    If Not RT_Confirm(api, "'" & rng.Cells(1, keyCol).Value2 & "' 열 기준 " & order.Count & "개 그룹으로 나눠 새 시트 " & order.Count & "개" & IIf(withSum, " + 소계 시트", "") & "를 만듭니다." & vbCrLf & "(원본은 바뀌지 않습니다.)") Then Exit Sub
    tok = api.CoGuard_Begin(True, True, False, True)
    On Error GoTo EH
    If withSum Then
        Set sumWs = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        sumWs.Name = Uniq(api, wb, "소계")
        sumWs.Cells(1, 1).Value2 = "그룹": sumWs.Cells(1, 2).Value2 = "줄 수"
        For j = 1 To rng.Columns.Count
            If j <> keyCol Then sumWs.Cells(1, sumWs.Cells(1, sumWs.Columns.Count).End(xlToLeft).Column + 1).Value2 = CStr(rng.Cells(1, j).Value2) & " 합계"
        Next j
        sumWs.Rows(1).Font.Bold = True
        sr = 1
    End If
    For Each grp In order
        Set rows1 = Nothing
        For i = 1 To groups(grp).Count
            If rows1 Is Nothing Then Set rows1 = rng.Rows(groups(grp)(i)) Else Set rows1 = Union(rows1, rng.Rows(groups(grp)(i)))
        Next i
        Set out = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        out.Name = Uniq(api, wb, CStr(grp))
        rng.Rows(1).Copy out.Cells(1, 1)
        rows1.Copy out.Cells(2, 1)
        Application.CutCopyMode = False
        n = groups(grp).Count
        out.Columns.AutoFit
        If firstOut Is Nothing Then Set firstOut = out
        made = made + 1
        If withSum Then
            sr = sr + 1
            sumWs.Cells(sr, 1).Value2 = CStr(grp)
            sumWs.Cells(sr, 2).Value2 = n
            c = 2
            For j = 1 To rng.Columns.Count
                If j <> keyCol Then
                    c = c + 1
                    sumWs.Cells(sr, c).Formula = "=SUM('" & Replace(out.Name, "'", "''") & "'!" & out.Range(out.Cells(2, j), out.Cells(n + 1, j)).Address(True, True) & ")"
                End If
            Next j
        End If
    Next grp
    If withSum Then sumWs.Columns.AutoFit
    api.CoGuard_End tok
    If withSum Then sumWs.Activate Else ws.Activate
    api.CoGuard_Notice made & "개 그룹 시트" & IIf(withSum, "와 소계 시트('" & sumWs.Name & "')", "") & "를 만들었습니다."
    Exit Sub
EH:
    msg = "쪼개는 중 문제가 생겼습니다(" & made & "개 만든 뒤): " & Err.Description
    Err.Clear
    On Error Resume Next
    Application.CutCopyMode = False
    api.CoGuard_End tok
    api.CoGuard_Notice msg
End Sub

Private Function Uniq(ByVal api As Object, ByVal wb As Workbook, ByVal wanted As String) As String
    Dim cand As String, n As Long, sh As Object, hit As Boolean, baseNm As String
    baseNm = api.CoText_SafeSheetName(wanted)
    cand = baseNm
    Do
        hit = False
        For Each sh In wb.Sheets
            If StrComp(sh.Name, cand, vbTextCompare) = 0 Then hit = True: Exit For
        Next sh
        If Not hit Then Exit Do
        n = n + 1
        cand = Left$(baseNm, 28) & "_" & (n + 1)
    Loop
    Uniq = cand
End Function
