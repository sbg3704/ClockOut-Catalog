Attribute VB_Name = "CoNum"
'==============================================================
' 기능: 칼퇴 공용 라이브러리 - 숫자·단위 도우미(순수 계산).
'       Office 개체를 쓰지 않아 Excel 없이도 같은 입력이면 같은 답이 나온다.
' 사용법:
'   CoNum_Clamp(값, 최소, 최대)          ' 설정값을 선언한 min/max 안으로 한 번 더 제한
'   CoNum_MmToPt(0.8) / CoNum_PtToMm / CoNum_CmToPt
'   If CoNum_TryParse("1.5", x) Then ...  ' 지역 설정과 무관하게 "."를 소수점으로 읽음
'   CoNum_RoundHalfUp(2.675, 2) -> 2.68  ' VBA Round(은행식 반올림)와 다른 보통 반올림
'   CoNum_SafeDiv(a, b, 0)               ' b가 0이면 대신 값
'   CoNum_TruncTo(1234, 100) -> 1200     ' 단위로 절사(0 쪽으로 버림). RoundHalfUp의 짝
'   CoNum_ToKoreanWon(1234)              ' -> "일금 일천이백삼십사원정" (False면 "일금 " 없이)
'   시험: Test_CoNum (직접 실행 창에 결과가 나온다)
' 의존: 없음
'==============================================================
Option Explicit
Option Private Module

Public Const CO_MM_TO_PT As Double = 72# / 25.4

Public Function CoNum_Clamp(ByVal v As Double, ByVal lo As Double, ByVal hi As Double) As Double
    If lo > hi Then CoNum_Clamp = v: Exit Function     ' 범위가 뒤집혀 있으면 손대지 않음
    If v < lo Then v = lo
    If v > hi Then v = hi
    CoNum_Clamp = v
End Function

Public Function CoNum_IsBetween(ByVal v As Double, ByVal lo As Double, ByVal hi As Double) As Boolean
    CoNum_IsBetween = (v >= lo And v <= hi)
End Function

Public Function CoNum_MmToPt(ByVal mm As Double) As Double
    CoNum_MmToPt = mm * CO_MM_TO_PT
End Function

Public Function CoNum_CmToPt(ByVal cm As Double) As Double
    CoNum_CmToPt = cm * 10# * CO_MM_TO_PT
End Function

Public Function CoNum_PtToMm(ByVal pt As Double) As Double
    CoNum_PtToMm = pt / CO_MM_TO_PT
End Function

Public Function CoNum_SafeDiv(ByVal a As Double, ByVal b As Double, Optional ByVal fallback As Double = 0) As Double
    If b = 0 Then CoNum_SafeDiv = fallback Else CoNum_SafeDiv = a / b
End Function

Public Function CoNum_Near(ByVal a As Double, ByVal b As Double, Optional ByVal eps As Double = 0.000000001) As Boolean
    CoNum_Near = (Abs(a - b) <= eps)
End Function

' 글자를 숫자로 읽는다. 형식: [공백][+|-]숫자[.숫자][e[+|-]숫자][공백]. 천 단위 쉼표는 받지 않는다.
' 지역 설정(소수점이 쉼표인 PC 등)과 무관하게 "."를 소수점으로 읽는다. 읽으면 True.
Public Function CoNum_TryParse(ByVal s As String, ByRef result As Double) As Boolean
    Dim i As Long, n As Long, c As String, digits As Long, expDigits As Long
    s = Trim$(s)
    n = Len(s)
    If n = 0 Then Exit Function
    i = 1
    c = Mid$(s, i, 1)
    If c = "+" Or c = "-" Then i = i + 1
    Do While i <= n
        c = Mid$(s, i, 1)
        If c < "0" Or c > "9" Then Exit Do
        digits = digits + 1: i = i + 1
    Loop
    If i <= n Then
        If Mid$(s, i, 1) = "." Then
            i = i + 1
            Do While i <= n
                c = Mid$(s, i, 1)
                If c < "0" Or c > "9" Then Exit Do
                digits = digits + 1: i = i + 1
            Loop
        End If
    End If
    If digits = 0 Then Exit Function
    If i <= n Then
        c = LCase$(Mid$(s, i, 1))
        If c <> "e" Then Exit Function
        i = i + 1
        If i <= n Then
            c = Mid$(s, i, 1)
            If c = "+" Or c = "-" Then i = i + 1
        End If
        Do While i <= n
            c = Mid$(s, i, 1)
            If c < "0" Or c > "9" Then Exit Function
            expDigits = expDigits + 1: i = i + 1
        Loop
        If expDigits = 0 Or expDigits > 3 Then Exit Function
    End If
    On Error GoTo Fail                                ' 1e999 같은 넘침
    result = Val(s)                                   ' Val은 항상 "."를 소수점으로 읽는다
    CoNum_TryParse = True
    Exit Function
Fail:
    result = 0
End Function

' 보통 반올림(0.5는 0에서 멀어지는 쪽). digits는 0~10. 아주 큰 수(1E15 넘음)는 그대로 돌려준다.
' Decimal로 계산해 2.675*100=267.4999… 같은 부동소수 오차를 피한다.
Public Function CoNum_RoundHalfUp(ByVal v As Double, Optional ByVal digits As Long = 0) As Double
    Dim f As Variant, d As Variant
    If digits < 0 Then digits = 0
    If digits > 10 Then digits = 10
    If Abs(v) > 1E+15 Then CoNum_RoundHalfUp = v: Exit Function
    f = CDec(10 ^ digits)
    d = Int(CDec(Abs(v)) * f + CDec(0.5)) / f
    CoNum_RoundHalfUp = Sgn(v) * CDbl(d)
End Function

' unitValue의 배수로 절사한다(0 쪽으로 버림). TruncTo(1234, 100) -> 1200, TruncTo(-1234, 100) -> -1200,
' TruncTo(1.239, 0.01) -> 1.23. unitValue가 0 이하이거나 아주 큰 수(1E15 넘음)는 그대로 돌려준다.
' CoNum_RoundHalfUp처럼 Decimal로 계산해 부동소수 오차를 줄인다(비교는 CoNum_Near로).
Public Function CoNum_TruncTo(ByVal v As Double, ByVal unitValue As Double) As Double
    Dim d As Variant
    CoNum_TruncTo = v
    If unitValue <= 0 Or Abs(v) > 1E+15 Then Exit Function
    On Error GoTo Fail
    d = Fix(CDec(v) / CDec(unitValue)) * CDec(unitValue)
    CoNum_TruncTo = CDbl(d)
    Exit Function
Fail:
    CoNum_TruncTo = v
End Function

' 금액을 한글로 쓴다: 12345678 -> "일금 일천이백삼십사만오천육백칠십팔원정".
'   - withPrefix=False면 앞의 "일금 "을 뺀다("일천이백삼십사원정"). 0 -> "영원정"("일금 영원정").
'   - 단위는 만·억·조. 십·백·천 앞에도 "일"을 붙인다(수표·계약서 표기).
'   - 음수·소수, 1경(10^16) 이상은 ""(호출 쪽이 처리, 오류를 던지지 않음).
Public Function CoNum_ToKoreanWon(ByVal amount As Double, Optional ByVal withPrefix As Boolean = True) As String
    Dim numText As String, body As String, part As String, n As Long, g As Long, units As Variant
    On Error GoTo Fail
    If amount < 0 Or amount <> Fix(amount) Or amount >= 1E+16 Then Exit Function
    numText = CStr(CDec(amount))
    If numText = "0" Then
        body = "영"
    Else
        units = Array("", "만", "억", "조")
        n = Len(numText)
        Do While n > 0
            If n >= 4 Then part = Mid$(numText, n - 3, 4) Else part = Left$(numText, n)
            part = KoreanGroup(part)
            If Len(part) > 0 Then body = part & units(g) & body
            n = n - 4
            g = g + 1
        Loop
    End If
    If withPrefix Then body = "일금 " & body
    CoNum_ToKoreanWon = body & "원정"
    Exit Function
Fail:
    CoNum_ToKoreanWon = ""
End Function

' 네 자리 이하 숫자 글자 -> "일천이백삼십사". 0000이면 "".
Private Function KoreanGroup(ByVal grp As String) As String
    Const DIGIT_NAMES As String = "영일이삼사오육칠팔구"
    Dim k As Long, dg As Long, outText As String, posNames As Variant
    posNames = Array("천", "백", "십", "")
    grp = Right$("0000" & grp, 4)
    For k = 1 To 4
        dg = Asc(Mid$(grp, k, 1)) - 48
        If dg > 0 Then outText = outText & Mid$(DIGIT_NAMES, dg + 1, 1) & posNames(k - 1)
    Next k
    KoreanGroup = outText
End Function

'--- 시험 (Excel 안에서 직접 실행: 보기 > 직접 실행 창) --------------------
Public Sub Test_CoNum()
    Dim f As Long, x As Double
    f = f + Chk("Clamp", CoNum_Clamp(25, 0, 20) = 20 And CoNum_Clamp(-1, 0, 20) = 0 And CoNum_Clamp(5, 0, 20) = 5)
    f = f + Chk("Clamp 뒤집힌 범위", CoNum_Clamp(5, 10, 0) = 5)
    f = f + Chk("MmToPt", CoNum_Near(CoNum_MmToPt(25.4), 72))
    f = f + Chk("CmToPt", CoNum_Near(CoNum_CmToPt(2.54), 72))
    f = f + Chk("PtToMm", CoNum_Near(CoNum_PtToMm(72), 25.4))
    f = f + Chk("SafeDiv", CoNum_SafeDiv(1, 0, -1) = -1 And CoNum_SafeDiv(6, 3) = 2)
    f = f + Chk("TryParse 소수", CoNum_TryParse(" 1.5 ", x) And CoNum_Near(x, 1.5))
    f = f + Chk("TryParse 음수", CoNum_TryParse("-2", x) And x = -2)
    f = f + Chk("TryParse 지수", CoNum_TryParse("1e3", x) And x = 1000)
    f = f + Chk("TryParse 점으로 시작", CoNum_TryParse(".5", x) And CoNum_Near(x, 0.5))
    f = f + Chk("TryParse 쉼표 거부", Not CoNum_TryParse("1,5", x))
    f = f + Chk("TryParse 글자 거부", Not CoNum_TryParse("12mm", x) And Not CoNum_TryParse("", x) And Not CoNum_TryParse("-", x))
    f = f + Chk("TryParse 지수 빈", Not CoNum_TryParse("1e", x))
    f = f + Chk("RoundHalfUp", CoNum_Near(CoNum_RoundHalfUp(2.5), 3) And CoNum_Near(CoNum_RoundHalfUp(3.5), 4))
    f = f + Chk("RoundHalfUp 오차", CoNum_Near(CoNum_RoundHalfUp(2.675, 2), 2.68))
    f = f + Chk("RoundHalfUp 음수", CoNum_Near(CoNum_RoundHalfUp(-2.5), -3))
    f = f + Chk("TruncTo", CoNum_Near(CoNum_TruncTo(1234, 100), 1200) And CoNum_Near(CoNum_TruncTo(99, 100), 0))
    f = f + Chk("TruncTo 음수", CoNum_Near(CoNum_TruncTo(-1234, 100), -1200))
    f = f + Chk("TruncTo 소수", CoNum_Near(CoNum_TruncTo(1.239, 0.01), 1.23) And CoNum_Near(CoNum_TruncTo(0.3, 0.1), 0.3))
    f = f + Chk("TruncTo 단위 0 이하", CoNum_TruncTo(12.5, 0) = 12.5 And CoNum_TruncTo(12.5, -1) = 12.5)
    f = f + Chk("ToKoreanWon 영", CoNum_ToKoreanWon(0, False) = "영원정" And CoNum_ToKoreanWon(0) = "일금 영원정")
    f = f + Chk("ToKoreanWon 천", CoNum_ToKoreanWon(1234) = "일금 일천이백삼십사원정")
    f = f + Chk("ToKoreanWon 만", CoNum_ToKoreanWon(12345678) = "일금 일천이백삼십사만오천육백칠십팔원정")
    f = f + Chk("ToKoreanWon 억", CoNum_ToKoreanWon(100000000) = "일금 일억원정")
    f = f + Chk("ToKoreanWon 빈 자리", CoNum_ToKoreanWon(100010, False) = "일십만일십원정")
    f = f + Chk("ToKoreanWon 조", CoNum_ToKoreanWon(2000000050000#, False) = "이조오만원정")
    f = f + Chk("ToKoreanWon 음수·소수", CoNum_ToKoreanWon(-1) = "" And CoNum_ToKoreanWon(1.5) = "")
    f = f + Chk("ToKoreanWon 너무 큼", CoNum_ToKoreanWon(1E+16) = "")
    Report "CoNum", f
End Sub

Private Function Chk(ByVal testName As String, ByVal ok As Boolean) As Long
    If Not ok Then Debug.Print "실패: " & testName: Chk = 1
End Function

Private Sub Report(ByVal moduleName As String, ByVal fails As Long)
    If fails = 0 Then
        Debug.Print moduleName & " 시험: 모두 통과"
        MsgBox moduleName & " 시험: 모두 통과", vbInformation, "칼퇴"
    Else
        MsgBox moduleName & " 시험: " & fails & "개 실패 (직접 실행 창 확인)", vbExclamation, "칼퇴"
    End If
End Sub
