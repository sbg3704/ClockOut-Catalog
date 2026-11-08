Attribute VB_Name = "RangeColor"
'==============================================================
' 기능: 선택한 칸을 칸 색(또는 글자색)별로 묶어 개수와 합계를 알려 준다. 엑셀 함수로는 색을 못 읽어 못 하던 일.
' 사용법: 범위를 고르고 우클릭 > 칼퇴 범위 도구 > 색깔별 합계.
' 설정: colorBy("칸 색" 또는 "글자색"), colorOutput("창으로 보기" 또는 "새 시트에 쓰기").
' 규칙: 값은 바꾸지 않는다(읽기만). 조건부 서식으로 보이는 색은 세지 않고 직접 칠한 색만 센다. 숫자가 아닌 칸은 개수에서 빼고 "글자 칸"으로 따로 센다.
' 의존: RangeTools.bas, 공용 라이브러리(CoSetting_Choice, CoGuard_*)
'==============================================================
Option Explicit
Option Private Module

Private Const MAX_CELLS As Long = 200000

Public Sub RangeColor_Sum()
    Dim api As Object, ws As Worksheet, rng As Range, byFont As Boolean, toSheet As Boolean
    Dim cel As Range, d As Object, key As Variant, col As Long, v As Variant, n As Long
    Dim sums As Object, nums As Object, texts As Object, tot As Double, txt As String, k As Long
    Set api = RT_Api()
    If api Is Nothing Then Exit Sub
    Set rng = RT_Area(api, ws)
    If rng Is Nothing Then Exit Sub
    If rng.Cells.CountLarge > MAX_CELLS Then api.CoGuard_Notice "칸이 너무 많습니다(최대 " & MAX_CELLS & "칸). 범위를 줄이세요.": Exit Sub
    byFont = (api.CoSetting_Choice(RT_ADDIN_ID, "colorBy", "칸 색", "칸 색", "글자색") = "글자색")
    toSheet = (api.CoSetting_Choice(RT_ADDIN_ID, "colorOutput", "창으로 보기", "창으로 보기", "새 시트에 쓰기") = "새 시트에 쓰기")
    Set sums = CreateObject("Scripting.Dictionary")
    Set nums = CreateObject("Scripting.Dictionary")
    Set texts = CreateObject("Scripting.Dictionary")
    For Each cel In rng.Cells
        If Not cel.EntireRow.Hidden And Not cel.EntireColumn.Hidden Then
            v = cel.Value2
            If Not IsEmpty(v) Then
                If byFont Then col = cel.Font.Color Else col = IIf(cel.Interior.ColorIndex = xlColorIndexNone, -1, cel.Interior.Color)
                If Not sums.Exists(col) Then sums.Add col, 0#: nums.Add col, 0: texts.Add col, 0
                If VarType(v) = vbDouble Then
                    sums(col) = sums(col) + v: nums(col) = nums(col) + 1
                ElseIf Not IsError(v) Then
                    texts(col) = texts(col) + 1
                End If
            End If
        End If
    Next cel
    If sums.Count = 0 Then api.CoGuard_Notice "값이 있는 칸이 없습니다.": Exit Sub
    If toSheet Then
        WriteSheet api, ws, sums, nums, texts, byFont
    Else
        For Each key In sums.Keys
            k = k + 1
            If k <= 15 Then txt = txt & Label(CLng(key), byFont) & ": 합계 " & Format$(sums(key), "#,##0.##") & "  (숫자 " & nums(key) & "칸" & _
                                  IIf(texts(key) > 0, ", 글자 " & texts(key) & "칸", "") & ")" & vbCrLf
            tot = tot + sums(key)
        Next key
        If k > 15 Then txt = txt & "... 외 " & (k - 15) & "가지 색" & vbCrLf
        api.CoGuard_Notice "[" & IIf(byFont, "글자색", "칸 색") & "별 합계] " & rng.Address(False, False) & vbCrLf & vbCrLf & txt & vbCrLf & "전체 합계: " & Format$(tot, "#,##0.##")
    End If
End Sub

Private Function Label(ByVal c As Long, ByVal byFont As Boolean) As String
    If c = -1 Then
        Label = "색 없음"
    Else
        Label = "RGB(" & (c Mod 256) & "," & ((c \ 256) Mod 256) & "," & (c \ 65536) & ")"
    End If
End Function

Private Sub WriteSheet(ByVal api As Object, ByVal ws As Worksheet, ByVal sums As Object, ByVal nums As Object, ByVal texts As Object, ByVal byFont As Boolean)
    Dim wb As Workbook, out As Worksheet, key As Variant, r As Long
    Set wb = ws.Parent
    If wb.ProtectStructure Then api.CoGuard_Notice "통합 문서 구조가 보호되어 있어 새 시트를 만들 수 없습니다.": Exit Sub
    Set out = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    out.Name = ST_Unique(api, wb, "색깔별합계")
    out.Range("A1:E1").Value2 = Array("색", "표시", "합계", "숫자 칸 수", "글자 칸 수")
    out.Range("A1:E1").Font.Bold = True
    r = 1
    For Each key In sums.Keys
        r = r + 1
        out.Cells(r, 1).Value2 = Label(CLng(key), byFont)
        If CLng(key) <> -1 Then
            If byFont Then out.Cells(r, 2).Font.Color = CLng(key): out.Cells(r, 2).Value2 = "가나다 ABC" Else out.Cells(r, 2).Interior.Color = CLng(key)
        End If
        out.Cells(r, 3).Value2 = sums(key)
        out.Cells(r, 4).Value2 = nums(key)
        out.Cells(r, 5).Value2 = texts(key)
    Next key
    out.Range("C2:C" & r).NumberFormat = "#,##0.##"
    out.Columns("A:E").AutoFit
    out.Activate
    api.CoGuard_Notice (r - 1) & "가지 색의 합계를 '" & out.Name & "' 시트에 정리했습니다."
End Sub

' 겹치지 않는 시트 이름.
Private Function ST_Unique(ByVal api As Object, ByVal wb As Workbook, ByVal wanted As String) As String
    Dim cand As String, n As Long, sh As Object, hit As Boolean
    cand = api.CoText_SafeSheetName(wanted)
    Do
        hit = False
        For Each sh In wb.Sheets
            If StrComp(sh.Name, cand, vbTextCompare) = 0 Then hit = True: Exit For
        Next sh
        If Not hit Then Exit Do
        n = n + 1
        cand = Left$(api.CoText_SafeSheetName(wanted), 28) & "_" & (n + 1)
    Loop
    ST_Unique = cand
End Function
