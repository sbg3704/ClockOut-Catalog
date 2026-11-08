Attribute VB_Name = "CoText"
'==============================================================
' 기능: 칼퇴 공용 라이브러리 - 글자·경로·열 이름 도우미(순수 계산).
'       Office 개체를 쓰지 않아 Excel 없이도 같은 입력이면 같은 답이 나온다.
' 사용법:
'   CoText_IsBlank(셀값)                   ' 빈칸·공백만·Null·Empty면 True
'   CoText_SafeFileName("보고서: 1/2")     ' -> "보고서_ 1_2" (파일 이름에 못 쓰는 글자 바꾸기)
'   CoText_PathJoin("C:\a", "b.txt")       ' -> "C:\a\b.txt"
'   CoText_FileName / CoText_BaseName / CoText_Ext / CoText_ParentFolder
'   CoText_ColLetter(28) -> "AB",  CoText_ColNumber("AB") -> 28
'   Set parts = CoText_SplitClean("a, b,,c", ",")   ' -> a, b, c (앞뒤 공백·빈 조각 버림)
'   CoText_CleanInvisible(" a" & ChrW(160) & " b ")    ' -> "a b" (NBSP·폭 0 글자·줄바꿈·겹친 공백 정리)
'   CoText_SafeSheetName("1/2분기[안]")      ' -> "1_2분기_안_" (시트 이름 규칙: 금지 문자·31자·앞뒤 ')
'   CoText_NaturalCompare("2호", "10호")      ' -> -1 (숫자는 크기로 비교, 대소문자 무시)
'   If CoText_ParseDate("2026.11.08", d) Then ...   ' 년월일 순서 고정, 지역 설정과 무관
'   CoText_JoinList(목록, ", ", "'", "'")    ' -> 'a', 'b'
'   시험: Test_CoText (직접 실행 창에 결과가 나온다)
' 의존: 없음
'==============================================================
Option Explicit
Option Private Module

' 비었는지(Empty, Null, 빈 글자, 공백만). 오류 값(#N/A 등)은 비지 않은 것으로 본다.
Public Function CoText_IsBlank(ByVal v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Then CoText_IsBlank = True: Exit Function
    If IsError(v) Or IsObject(v) Or IsArray(v) Then Exit Function
    CoText_IsBlank = (Len(Trim$(CStr(v))) = 0)
End Function

Public Function CoText_StartsWith(ByVal s As String, ByVal prefix As String, _
                                  Optional ByVal ignoreCase As Boolean = False) As Boolean
    If Len(prefix) > Len(s) Then Exit Function
    If ignoreCase Then
        CoText_StartsWith = (StrComp(Left$(s, Len(prefix)), prefix, vbTextCompare) = 0)
    Else
        CoText_StartsWith = (Left$(s, Len(prefix)) = prefix)
    End If
End Function

Public Function CoText_EndsWith(ByVal s As String, ByVal suffix As String, _
                                Optional ByVal ignoreCase As Boolean = False) As Boolean
    If Len(suffix) > Len(s) Then Exit Function
    If ignoreCase Then
        CoText_EndsWith = (StrComp(Right$(s, Len(suffix)), suffix, vbTextCompare) = 0)
    Else
        CoText_EndsWith = (Right$(s, Len(suffix)) = suffix)
    End If
End Function

' 조각으로 나눠 앞뒤 공백을 지우고 빈 조각은 버린다. String들의 Collection.
Public Function CoText_SplitClean(ByVal s As String, Optional ByVal delim As String = ",") As Collection
    Dim col As New Collection, parts() As String, i As Long, t As String
    If Len(s) > 0 And Len(delim) > 0 Then
        parts = Split(s, delim)
        For i = LBound(parts) To UBound(parts)
            t = Trim$(parts(i))
            If Len(t) > 0 Then col.Add t
        Next i
    End If
    Set CoText_SplitClean = col
End Function

' Windows 파일 이름에 못 쓰는 글자(\ / : * ? " < > | 와 제어 문자)를 바꾸고,
' 끝의 점·공백을 지우고, 예약 이름(CON, NUL, COM1 등)이면 앞에 _를 붙인다. 비면 "_".
Public Function CoText_SafeFileName(ByVal fileName As String, Optional ByVal replaceWith As String = "_") As String
    Dim i As Long, c As String, outText As String, stem As String, p As Long
    For i = 1 To Len(fileName)
        c = Mid$(fileName, i, 1)
        If InStr(1, "\/:*?""<>|", c, vbBinaryCompare) > 0 Or AscW(c) >= 0 And AscW(c) < 32 Then
            outText = outText & replaceWith
        Else
            outText = outText & c
        End If
    Next i
    Do While Len(outText) > 0 And (Right$(outText, 1) = "." Or Right$(outText, 1) = " ")
        outText = Left$(outText, Len(outText) - 1)
    Loop
    If Len(outText) = 0 Then CoText_SafeFileName = "_": Exit Function
    p = InStr(1, outText, ".")
    If p > 0 Then stem = Left$(outText, p - 1) Else stem = outText
    Select Case UCase$(Trim$(stem))
        Case "CON", "PRN", "AUX", "NUL", _
             "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", _
             "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
            outText = "_" & outText
    End Select
    CoText_SafeFileName = outText
End Function

' 경로 끝의 \ 또는 /를 지운다. 드라이브 뿌리("C:\")는 그대로 둔다.
Public Function CoText_TrimSlash(ByVal path As String) As String
    Do While Len(path) > 0 And (Right$(path, 1) = "\" Or Right$(path, 1) = "/")
        If Len(path) = 3 And Mid$(path, 2, 1) = ":" Then Exit Do
        path = Left$(path, Len(path) - 1)
    Loop
    CoText_TrimSlash = path
End Function

' 폴더와 이름을 \ 하나로 잇는다. name이 비면 폴더 끝에 \를 붙여 돌려준다.
Public Function CoText_PathJoin(ByVal folder As String, ByVal child As String) As String
    Dim f As String
    f = CoText_TrimSlash(folder)
    Do While Len(child) > 0 And (Left$(child, 1) = "\" Or Left$(child, 1) = "/")
        child = Mid$(child, 2)
    Loop
    If Len(f) = 0 Then CoText_PathJoin = child: Exit Function
    If Right$(f, 1) = "\" Then
        CoText_PathJoin = f & child                    ' "C:\" 같은 드라이브 뿌리
    Else
        CoText_PathJoin = f & "\" & child
    End If
End Function

Private Function LastSlash(ByVal path As String) As Long
    Dim a As Long, b As Long
    a = InStrRev(path, "\")
    b = InStrRev(path, "/")
    If b > a Then LastSlash = b Else LastSlash = a
End Function

' "C:\a\b.txt" -> "b.txt"
Public Function CoText_FileName(ByVal path As String) As String
    CoText_FileName = Mid$(path, LastSlash(path) + 1)
End Function

' "C:\a\b.tar.gz" -> "gz" (점 없이). 확장자가 없거나 ".gitignore"처럼 점으로 시작하면 "".
Public Function CoText_Ext(ByVal path As String) As String
    Dim n As String, p As Long
    n = CoText_FileName(path)
    p = InStrRev(n, ".")
    If p > 1 Then CoText_Ext = Mid$(n, p + 1)
End Function

' "C:\a\b.txt" -> "b"
Public Function CoText_BaseName(ByVal path As String) As String
    Dim n As String, p As Long
    n = CoText_FileName(path)
    p = InStrRev(n, ".")
    If p > 1 Then CoText_BaseName = Left$(n, p - 1) Else CoText_BaseName = n
End Function

' "C:\a\b.txt" -> "C:\a",  "C:\a" -> "C:\",  "C:\" -> "",  "b.txt" -> ""
Public Function CoText_ParentFolder(ByVal path As String) As String
    Dim p As String, k As Long
    p = CoText_TrimSlash(path)
    If Len(p) = 3 And Mid$(p, 2, 2) = ":\" Then Exit Function
    k = LastSlash(p)
    If k = 0 Then Exit Function
    If k = 3 And Mid$(p, 2, 1) = ":" Then
        CoText_ParentFolder = Left$(p, 3)            ' 드라이브 뿌리는 \까지
    Else
        CoText_ParentFolder = Left$(p, k - 1)
    End If
End Function

' 열 번호 -> 열 이름(1 -> "A", 28 -> "AB"). 1~16384 밖이면 "".
Public Function CoText_ColLetter(ByVal n As Long) As String
    Dim s As String, r As Long
    If n < 1 Or n > 16384 Then Exit Function
    Do While n > 0
        r = (n - 1) Mod 26
        s = Chr$(65 + r) & s
        n = (n - 1) \ 26
    Loop
    CoText_ColLetter = s
End Function

' 열 이름 -> 열 번호("AB" -> 28, 대소문자 무관). 잘못된 이름이거나 16384 넘으면 0.
Public Function CoText_ColNumber(ByVal s As String) As Long
    Dim i As Long, c As Long, n As Long
    s = UCase$(Trim$(s))
    If Len(s) = 0 Or Len(s) > 3 Then Exit Function
    For i = 1 To Len(s)
        c = Asc(Mid$(s, i, 1))
        If c < 65 Or c > 90 Then Exit Function
        n = n * 26 + (c - 64)
    Next i
    If n > 16384 Then Exit Function
    CoText_ColNumber = n
End Function

' 보이지 않는 글자를 정리한다.
'   - NBSP(160)·전각 공백·폭 있는 각종 공백(U+2000~200A 등)·탭 -> 보통 공백
'   - 폭 0 글자(ZWSP·ZWJ·BOM·단어 결합자)·소프트 하이픈·제어 문자 -> 지움
'   - 줄바꿈(CR·LF): keepLineBreak=False면 공백, True면 줄바꿈 하나(vbLf, 셀 안 줄바꿈과 같음)로 남기고
'     줄마다 앞뒤 공백을 지우며 빈 줄은 없앤다
'   - 공백 여러 개는 하나로, 앞뒤 공백은 지운다
Public Function CoText_CleanInvisible(ByVal s As String, Optional ByVal keepLineBreak As Boolean = False) As String
    Dim i As Long, code As Long, c As String, buf As String, lines() As String, outText As String, ln As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        code = AscW(c) And &HFFFF&
        Select Case code
            Case 10, 13
                If keepLineBreak Then buf = buf & vbLf Else buf = buf & " "
            Case 9, 160, 5760, 8192 To 8202, 8239, 8287, 12288
                buf = buf & " "
            Case 0 To 31, 127, 173, 8203 To 8205, 8288, 65279
                ' 지움
            Case Else
                buf = buf & c
        End Select
    Next i
    If Not keepLineBreak Then CoText_CleanInvisible = CollapseSpaces(buf): Exit Function
    lines = Split(buf, vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = CollapseSpaces(lines(i))
        If Len(ln) > 0 Then
            If Len(outText) > 0 Then outText = outText & vbLf
            outText = outText & ln
        End If
    Next i
    CoText_CleanInvisible = outText
End Function

Private Function CollapseSpaces(ByVal t As String) As String
    Do While InStr(1, t, "  ", vbBinaryCompare) > 0
        t = Replace(t, "  ", " ")
    Loop
    CollapseSpaces = Trim$(t)
End Function

' 시트 이름 규칙에 맞춘 이름: 금지 문자(: \ / ? * [ ])는 replaceWith로 바꾸고, 앞뒤 '는 지우고,
' 31자로 자른다. 비면 "Sheet". Excel이 쓰는 예약 이름 "History"는 뒤에 replaceWith를 붙인다.
' 같은 이름의 시트가 이미 있는지는 보지 않는다(통합 문서를 봐야 하므로 호출 쪽이 확인).
Public Function CoText_SafeSheetName(ByVal sheetName As String, Optional ByVal replaceWith As String = "_") As String
    Const BAD_CHARS As String = ":\/?*[]"
    Dim i As Long, c As String, outText As String
    If Len(replaceWith) > 0 Then                      ' 바꿀 글자에 금지 문자가 섞여 있으면 _로
        For i = 1 To Len(replaceWith)
            If InStr(1, BAD_CHARS, Mid$(replaceWith, i, 1), vbBinaryCompare) > 0 Then replaceWith = "_": Exit For
        Next i
    End If
    For i = 1 To Len(sheetName)
        c = Mid$(sheetName, i, 1)
        If InStr(1, BAD_CHARS, c, vbBinaryCompare) > 0 Then outText = outText & replaceWith Else outText = outText & c
    Next i
    outText = TrimQuotes(outText)
    If Len(outText) > 31 Then outText = TrimQuotes(Left$(outText, 31))
    If Len(outText) = 0 Then outText = "Sheet"
    If StrComp(outText, "History", vbTextCompare) = 0 Then outText = Left$(outText & replaceWith, 31)
    CoText_SafeSheetName = outText
End Function

Private Function TrimQuotes(ByVal t As String) As String
    Do While Left$(t, 1) = "'"
        t = Mid$(t, 2)
    Loop
    Do While Right$(t, 1) = "'"
        t = Left$(t, Len(t) - 1)
    Loop
    TrimQuotes = t
End Function

' 자연 순서 비교: 숫자 부분은 크기로, 나머지는 대소문자 무시로 비교한다. a<b -> -1, 같으면 0, a>b -> 1.
' "1호" < "2호" < "10호", "a2" < "a10". 앞의 0은 무시한다("007" = "7").
Public Function CoText_NaturalCompare(ByVal a As String, ByVal b As String) As Long
    Dim i As Long, j As Long, ra As String, rb As String, r As Long
    i = 1: j = 1
    Do While i <= Len(a) And j <= Len(b)
        ra = TakeRun(a, i)
        rb = TakeRun(b, j)
        If IsDigitChar(Left$(ra, 1)) And IsDigitChar(Left$(rb, 1)) Then
            r = CompareDigitRuns(ra, rb)
        Else
            r = StrComp(ra, rb, vbTextCompare)
        End If
        If r <> 0 Then CoText_NaturalCompare = r: Exit Function
    Loop
    If i <= Len(a) Then
        CoText_NaturalCompare = 1                     ' a가 더 김
    ElseIf j <= Len(b) Then
        CoText_NaturalCompare = -1
    End If
End Function

Private Function IsDigitChar(ByVal c As String) As Boolean
    IsDigitChar = (Len(c) = 1 And c >= "0" And c <= "9")
End Function

' i부터 같은 종류(숫자 / 숫자 아님)가 이어지는 조각을 돌려주고 i를 그 다음으로 옮긴다.
Private Function TakeRun(ByVal s As String, ByRef i As Long) As String
    Dim start As Long, digitRun As Boolean
    start = i
    digitRun = IsDigitChar(Mid$(s, i, 1))
    Do While i <= Len(s)
        If IsDigitChar(Mid$(s, i, 1)) <> digitRun Then Exit Do
        i = i + 1
    Loop
    TakeRun = Mid$(s, start, i - start)
End Function

Private Function CompareDigitRuns(ByVal x As String, ByVal y As String) As Long
    Do While Len(x) > 1 And Left$(x, 1) = "0"
        x = Mid$(x, 2)
    Loop
    Do While Len(y) > 1 And Left$(y, 1) = "0"
        y = Mid$(y, 2)
    Loop
    If Len(x) <> Len(y) Then
        If Len(x) < Len(y) Then CompareDigitRuns = -1 Else CompareDigitRuns = 1
    Else
        CompareDigitRuns = StrComp(x, y, vbBinaryCompare)
    End If
End Function

' 글자를 날짜로 읽는다. 년월일 순서 고정(지역 설정과 무관). 읽으면 True, 못 읽으면 False(result는 0).
'   받는 꼴: "2026.11.08", "2026-11-08", "2026/11/08", "2026. 11. 8."(띄어쓰기·끝 점), "20261108",
'            두 자리 해 "26/11/08"(2000년대로 본다). 없는 날("2026.13.40", "2026.02.29")은 False.
Public Function CoText_ParseDate(ByVal s As String, ByRef result As Date) As Boolean
    Dim t As String, parts() As String, y As Long, m As Long, d As Long, k As Long
    result = 0
    t = Replace(Trim$(s), " ", "")
    If Len(t) = 0 Then Exit Function
    If Right$(t, 1) = "." Then t = Left$(t, Len(t) - 1)
    If Len(t) = 8 And AllDigits(t) Then
        y = CLng(Left$(t, 4)): m = CLng(Mid$(t, 5, 2)): d = CLng(Right$(t, 2))
    Else
        t = Replace(Replace(t, "-", "."), "/", ".")
        parts = Split(t, ".")
        If UBound(parts) - LBound(parts) <> 2 Then Exit Function
        For k = LBound(parts) To UBound(parts)
            If Len(parts(k)) = 0 Or Not AllDigits(parts(k)) Then Exit Function
        Next k
        Select Case Len(parts(0))
            Case 4: y = CLng(parts(0))
            Case 2: y = 2000 + CLng(parts(0))
            Case Else: Exit Function
        End Select
        If Len(parts(1)) > 2 Or Len(parts(2)) > 2 Then Exit Function
        m = CLng(parts(1)): d = CLng(parts(2))
    End If
    If y < 1000 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
    result = DateSerial(y, m, d)
    If Year(result) <> y Or Month(result) <> m Or Day(result) <> d Then result = 0: Exit Function   ' 2월 30일 등
    CoText_ParseDate = True
End Function

Private Function AllDigits(ByVal t As String) As Boolean
    Dim i As Long
    If Len(t) = 0 Then Exit Function
    For i = 1 To Len(t)
        If Not IsDigitChar(Mid$(t, i, 1)) Then Exit Function
    Next i
    AllDigits = True
End Function

' 목록의 값마다 prefix·suffix를 붙여 delim으로 잇는다. 빈 목록이나 Nothing이면 "".
' 개체(셀 등)는 건너뛰고, Null·Empty는 빈 글자로 본다. 숫자는 CStr이라 소수점은 PC 지역 설정을 따른다.
Public Function CoText_JoinList(ByVal items As Collection, ByVal delim As String, _
                                Optional ByVal prefix As String = "", Optional ByVal suffix As String = "") As String
    Dim v As Variant, outText As String, first As Boolean, piece As String
    If items Is Nothing Then Exit Function
    first = True
    For Each v In items
        If Not IsObject(v) Then
            If IsNull(v) Or IsEmpty(v) Or IsError(v) Then piece = "" Else piece = CStr(v)
            If Not first Then outText = outText & delim
            outText = outText & prefix & piece & suffix
            first = False
        End If
    Next v
    CoText_JoinList = outText
End Function

'--- 시험 (Excel 안에서 직접 실행: 보기 > 직접 실행 창) --------------------
Public Sub Test_CoText()
    Dim f As Long, parts As Collection
    f = f + Chk("IsBlank 공백", CoText_IsBlank("   "))
    f = f + Chk("IsBlank Empty", CoText_IsBlank(Empty))
    f = f + Chk("IsBlank 0", Not CoText_IsBlank(0))
    f = f + Chk("IsBlank 오류값", Not CoText_IsBlank(CVErr(2042)))
    f = f + Chk("StartsWith", CoText_StartsWith("ClockOut", "clock", True) And Not CoText_StartsWith("ClockOut", "clock"))
    f = f + Chk("EndsWith", CoText_EndsWith("a.XLAM", ".xlam", True) And Not CoText_EndsWith("a", "abc"))
    Set parts = CoText_SplitClean(" a, b,,c ,", ",")
    f = f + Chk("SplitClean", parts.Count = 3 And parts(1) = "a" And parts(3) = "c")
    f = f + Chk("SplitClean 빈 글", CoText_SplitClean("", ",").Count = 0)
    f = f + Chk("SafeFileName 금지 글자", CoText_SafeFileName("보고서: 1/2?") = "보고서_ 1_2_")
    f = f + Chk("SafeFileName 끝 점", CoText_SafeFileName("이름. . ") = "이름")
    f = f + Chk("SafeFileName 예약", CoText_SafeFileName("con.txt") = "_con.txt")
    f = f + Chk("SafeFileName 빔", CoText_SafeFileName("...") = "_")
    f = f + Chk("TrimSlash", CoText_TrimSlash("C:\a\\") = "C:\a" And CoText_TrimSlash("C:\") = "C:\")
    f = f + Chk("PathJoin", CoText_PathJoin("C:\a\", "\b.txt") = "C:\a\b.txt")
    f = f + Chk("PathJoin 뿌리", CoText_PathJoin("C:\", "b") = "C:\b")
    f = f + Chk("PathJoin 폴더만", CoText_PathJoin("C:\a", "") = "C:\a\")
    f = f + Chk("FileName", CoText_FileName("C:\a\b.txt") = "b.txt" And CoText_FileName("b.txt") = "b.txt")
    f = f + Chk("Ext", CoText_Ext("C:\a.b\c.tar.gz") = "gz" And CoText_Ext("C:\a.b\c") = "" And CoText_Ext(".gitignore") = "")
    f = f + Chk("BaseName", CoText_BaseName("C:\a\b.txt") = "b" And CoText_BaseName("C:\a.b\c") = "c")
    f = f + Chk("ParentFolder", CoText_ParentFolder("C:\a\b.txt") = "C:\a")
    f = f + Chk("ParentFolder 뿌리", CoText_ParentFolder("C:\a") = "C:\" And CoText_ParentFolder("C:\") = "")
    f = f + Chk("ParentFolder 이름만", CoText_ParentFolder("b.txt") = "")
    f = f + Chk("ColLetter", CoText_ColLetter(1) = "A" And CoText_ColLetter(26) = "Z" And CoText_ColLetter(28) = "AB" And CoText_ColLetter(16384) = "XFD")
    f = f + Chk("ColLetter 범위 밖", CoText_ColLetter(0) = "" And CoText_ColLetter(16385) = "")
    f = f + Chk("ColNumber", CoText_ColNumber("ab") = 28 And CoText_ColNumber("XFD") = 16384)
    f = f + Chk("ColNumber 잘못", CoText_ColNumber("A1") = 0 And CoText_ColNumber("XFE") = 0 And CoText_ColNumber("") = 0)
    f = f + Chk("CleanInvisible 공백", CoText_CleanInvisible("  a   b  ") = "a b")
    f = f + Chk("CleanInvisible NBSP·전각", CoText_CleanInvisible("a" & ChrW(160) & "b" & ChrW(12288) & "c") = "a b c")
    f = f + Chk("CleanInvisible 폭 0", CoText_CleanInvisible(ChrW(65279) & "a" & ChrW(8203) & "b" & ChrW(173)) = "ab")
    f = f + Chk("CleanInvisible 줄바꿈", CoText_CleanInvisible("a" & vbCrLf & vbTab & "b") = "a b")
    f = f + Chk("CleanInvisible 줄 유지", CoText_CleanInvisible(" a " & vbCrLf & vbCrLf & " b  c" & vbCr, True) = "a" & vbLf & "b c")
    f = f + Chk("CleanInvisible 빔", CoText_CleanInvisible(vbTab & ChrW(160) & vbLf) = "")
    f = f + Chk("CleanInvisible 한글", CoText_CleanInvisible(" 칼퇴" & ChrW(160) & "매니저 ") = "칼퇴 매니저")
    Dim dt As Date, lst As Collection
    f = f + Chk("SafeSheetName 금지 문자", CoText_SafeSheetName("1/2분기[안]") = "1_2분기_안_" And CoText_SafeSheetName("a:b?c*d\e", "-") = "a-b-c-d-e")
    f = f + Chk("SafeSheetName 31자", Len(CoText_SafeSheetName(String$(40, "가"))) = 31)
    f = f + Chk("SafeSheetName 따옴표·빔", CoText_SafeSheetName("'보고서'") = "보고서" And CoText_SafeSheetName("''") = "Sheet" And CoText_SafeSheetName("") = "Sheet")
    f = f + Chk("SafeSheetName 예약", CoText_SafeSheetName("history") = "history_")
    f = f + Chk("NaturalCompare 호", CoText_NaturalCompare("1호", "2호") = -1 And CoText_NaturalCompare("2호", "10호") = -1 And CoText_NaturalCompare("10호", "2호") = 1)
    f = f + Chk("NaturalCompare 영문", CoText_NaturalCompare("a2", "a10") = -1 And CoText_NaturalCompare("A1", "a1") = 0)
    f = f + Chk("NaturalCompare 길이·0", CoText_NaturalCompare("a", "a1") = -1 And CoText_NaturalCompare("007", "7") = 0 And CoText_NaturalCompare("", "") = 0)
    f = f + Chk("ParseDate 점", CoText_ParseDate("2026.11.08", dt) And dt = DateSerial(2026, 11, 8))
    f = f + Chk("ParseDate 하이픈", CoText_ParseDate("2026-11-08", dt) And dt = DateSerial(2026, 11, 8))
    f = f + Chk("ParseDate 8자리", CoText_ParseDate("20261108", dt) And dt = DateSerial(2026, 11, 8))
    f = f + Chk("ParseDate 두 자리 해", CoText_ParseDate("26/11/08", dt) And dt = DateSerial(2026, 11, 8))
    f = f + Chk("ParseDate 띄어쓰기", CoText_ParseDate("2026. 11. 8.", dt) And dt = DateSerial(2026, 11, 8))
    f = f + Chk("ParseDate 없는 날", Not CoText_ParseDate("2026.13.40", dt) And Not CoText_ParseDate("2026.02.29", dt) And CoText_ParseDate("2024.2.29", dt))
    f = f + Chk("ParseDate 잘못", Not CoText_ParseDate("", dt) And Not CoText_ParseDate("2026.11", dt) And Not CoText_ParseDate("abc", dt))
    Set lst = New Collection
    f = f + Chk("JoinList 빔", CoText_JoinList(lst, ", ") = "" And CoText_JoinList(Nothing, ", ") = "")
    lst.Add "a": lst.Add "b"
    f = f + Chk("JoinList", CoText_JoinList(lst, ", ") = "a, b")
    f = f + Chk("JoinList 따옴표", CoText_JoinList(lst, ", ", "'", "'") = "'a', 'b'")
    Report "CoText", f
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
