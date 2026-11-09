Attribute VB_Name = "SelStatsCalc"
'==============================================================
' 기능: 선택한 칸들의 개수·숫자 개수·합계·평균·최소·최대를 계산하고 글자로 꾸민다.
'       엑셀 아래 상태 표시줄의 평균·개수·합계와 같은 값이다(숨긴 줄·오류 칸은 뺀다).
' 사용법: SelStatsCalc_Run(선택 범위, 결과) -> True면 결과를 쓴다. SelStatsCalc_Fmt(값, 천 단위 쉼표 여부)로 글자 만들기.
' 규칙: 선택이 아주 커도 느려지지 않게 "사용 중인 범위"와 겹치는 부분만 계산한다.
' 시험: Test_SelStatsCalc (글자 꾸미기 + 새 통합 문서에서 상태 표시줄과 같은 값인지, 저장하지 않고 닫는다)
' 의존: 없음
'==============================================================
Option Explicit
Option Private Module

Public Type SelStatsResult
    Count As Long          ' 비어 있지 않은 칸 수(상태 표시줄의 "개수")
    NumCount As Long       ' 숫자인 칸 수
    Sum As Double
    Avg As Double
    Min As Double
    Max As Double
End Type

Public Function SelStatsCalc_Run(ByVal sel As Range, ByRef r As SelStatsResult) As Boolean
    Dim ws As Worksheet, ar As Range, a As Range, used As Range, init As SelStatsResult
    r = init
    On Error GoTo EH
    Set ws = sel.Worksheet
    Set used = ws.UsedRange
    For Each ar In sel.Areas
        Set a = Intersect(ar, used)
        If Not a Is Nothing Then AddArea a, r
    Next ar
    If r.NumCount > 0 Then r.Avg = r.Sum / r.NumCount
    SelStatsCalc_Run = True
    Exit Function
EH:
    SelStatsCalc_Run = False
End Function

' 한 덩어리를 더한다. Aggregate 옵션 7 = 숨긴 줄과 오류 값을 무시.
Private Sub AddArea(ByVal a As Range, ByRef r As SelStatsResult)
    Dim wf As Object, n As Long, v As Double
    Set wf = Application.WorksheetFunction
    r.Count = r.Count + CLng(wf.Aggregate(3, 7, a))
    n = CLng(wf.Aggregate(2, 7, a))
    If n = 0 Then Exit Sub
    r.Sum = r.Sum + wf.Aggregate(9, 7, a)
    v = wf.Aggregate(5, 7, a)
    If r.NumCount = 0 Or v < r.Min Then r.Min = v
    v = wf.Aggregate(4, 7, a)
    If r.NumCount = 0 Or v > r.Max Then r.Max = v
    r.NumCount = r.NumCount + n
End Sub

' 값을 보기 좋은 글자로. 부동소수 잡음(0.30000000000000004)은 소수 10자리에서 정리한다.
Public Function SelStatsCalc_Fmt(ByVal v As Double, ByVal thousands As Boolean) As String
    Dim s As String
    If thousands Then s = Format$(Round(v, 10), "#,##0.##########") Else s = Format$(Round(v, 10), "0.##########")
    If Right$(s, 1) = "." Or Right$(s, 1) = "," Then s = Left$(s, Len(s) - 1)
    If s = "-0" Then s = "0"
    SelStatsCalc_Fmt = s
End Function

'--- 시험 (직접 실행 창에서 실행) ----------------------------------
Public Sub Test_SelStatsCalc()
    Dim f As Long
    f = f + Chk("정수", SelStatsCalc_Fmt(15, False) = "15")
    f = f + Chk("쉼표", SelStatsCalc_Fmt(1234567, True) = "1,234,567")
    f = f + Chk("쉼표 없음", SelStatsCalc_Fmt(1234567, False) = "1234567")
    f = f + Chk("소수", SelStatsCalc_Fmt(2.5, False) = "2.5")
    f = f + Chk("부동소수 잡음", SelStatsCalc_Fmt(0.1 + 0.2, False) = "0.3")
    f = f + Chk("음수", SelStatsCalc_Fmt(-1234.5, True) = "-1,234.5")
    f = f + Chk("0", SelStatsCalc_Fmt(0, True) = "0")
    f = f + SheetTest()
    If f = 0 Then
        Debug.Print "SelStatsCalc 시험: 모두 통과"
        MsgBox "SelStatsCalc 시험: 모두 통과", vbInformation, "칼퇴"
    Else
        MsgBox "SelStatsCalc 시험: " & f & "개 실패 (직접 실행 창 확인)", vbExclamation, "칼퇴"
    End If
End Sub

' 새 통합 문서에 상태 표시줄 예시(1,2,3,빈칸,4,빈칸,5)를 넣어 평균 3·개수 5·합계 15가 나오는지, 숨긴 줄·오류 칸이 빠지는지 본다.
Private Function SheetTest() As Long
    Dim wb As Workbook, ws As Worksheet, r As SelStatsResult, f As Long, ok As Boolean
    Set wb = Workbooks.Add
    On Error GoTo EH
    Set ws = wb.Worksheets(1)
    ws.Range("G2").Value = 1: ws.Range("G3").Value = 2: ws.Range("G4").Value = 3
    ws.Range("G6").Value = 4: ws.Range("G8").Value = 5
    ok = SelStatsCalc_Run(ws.Range("G2:G9"), r)
    f = f + Chk("예시 계산", ok And r.Count = 5 And r.NumCount = 5 And r.Sum = 15 And r.Avg = 3 And r.Min = 1 And r.Max = 5)
    ws.Range("G10").Formula = "=1/0"
    ok = SelStatsCalc_Run(ws.Range("G2:G10"), r)
    f = f + Chk("오류 칸 제외", ok And r.NumCount = 5 And r.Sum = 15)
    ws.Rows(8).Hidden = True
    ok = SelStatsCalc_Run(ws.Range("G2:G9"), r)
    f = f + Chk("숨긴 줄 제외", ok And r.NumCount = 4 And r.Sum = 10 And r.Max = 4)
    ok = SelStatsCalc_Run(ws.Range("A1:A3,G2:G4"), r)
    f = f + Chk("떨어진 선택", ok And r.NumCount = 3 And r.Sum = 6)
    ok = SelStatsCalc_Run(ws.Range("A1:A3"), r)
    f = f + Chk("빈 선택", ok And r.Count = 0 And r.NumCount = 0)
    ok = SelStatsCalc_Run(ws.Cells, r)
    f = f + Chk("시트 전체 선택", ok And r.NumCount = 4)
    wb.Close False
    SheetTest = f
    Exit Function
EH:
    Debug.Print "[실패] 시트 시험 오류: " & Err.Description
    On Error Resume Next
    wb.Close False
    SheetTest = f + 1
End Function

Private Function Chk(ByVal testName As String, ByVal ok As Boolean) As Long
    If Not ok Then Debug.Print "[실패] " & testName: Chk = 1
End Function
