Attribute VB_Name = "CoKey"
'==============================================================
' 기능: 칼퇴 공용 라이브러리 - 단축키 등록부. 여러 추가기능이 같은 키를 겹쳐 잡거나 엑셀 기본 키(Ctrl+C 등)를
'       덮어쓰지 않게, Application.OnKey 등록을 이 모듈 한 곳에서 한다(요청 004).
' 사용법:
'   If Not CoKey_Register("Ctrl+Shift+H", "'" & ThisWorkbook.Name & "'!CrossHair_Toggle", ADDIN_ID, why) Then Debug.Print why
'   CoKey_Unregister "Ctrl+Shift+H", ADDIN_ID           ' 등록한 소유자만 풀 수 있다
'   CoKey_UnregisterAll ADDIN_ID                         ' Workbook_BeforeClose에서 그 추가기능 키를 모두 풀기
'   If CoKey_Conflicts("ctrl + shift + h", owners) Then  ' 등록하지 않고 누가 쓰는지만 본다
'   s = CoKey_Normalize("shift+ctrl+h")   -> "Ctrl+Shift+H"  (형식이 틀리면 "")
'   s = CoKey_ToOnKey("Ctrl+Shift+H")     -> "^+H"           (OnKey 형식, 앱 저장값 key.<id>와 같은 정규형)
'   If CoKey_IsDangerous("Ctrl+V", why) Then ...          ' 엑셀 기본 키·글자 한 개 키
' 키 표기: "Ctrl+Shift+H"(대소문자·순서·공백 무시)와 OnKey 형식 "^+h", "{F5}" 둘 다 받는다.
'   정규형은 수정키 순서 Ctrl→Shift→Alt, 글자는 대문자(앱 ClockOut.Core.Shortcuts.ShortcutKey와 같은 규칙).
' 등록부: 이 모듈의 Collection 하나. clockout-lib.xlam은 Excel 한 실행에 하나만 열리므로 모든 추가기능이 공유한다.
'   (lib를 bundled로 함께 컴파일한 추가기능은 자기 사본의 등록부를 따로 가진다. 독립 lib를 쓰도록 권한다.)
' 시험: Test_CoKey(CoKeyTest.bas). 정규화·위험 키는 순수 계산, 등록은 Ctrl+Shift+Alt+F15로 Excel 안에서 확인
' 의존: 없음
'==============================================================
Option Explicit
Option Private Module

' 위험 키(OnKey 정규형). 확실히 위험한 것만 이 한 곳에 둔다: 엑셀 기본 편집·파일·서식 키, 되돌리기·붙여넣기 계열, F1~F12 단독, 창 닫기.
Private Const DANGER_KEYS As String = _
    "^C ^V ^X ^Z ^Y ^S ^A ^F ^H ^P ^N ^O ^W ^B ^I ^U ^G ^K ^D ^R ^+Z ^+V ^%V " & _
    "{F1} {F2} {F3} {F4} {F5} {F6} {F7} {F8} {F9} {F10} {F11} {F12} %{F4} ^{F4} ^{F6} ^{F9} ^{F10}"
' 이름 키: 표시 이름 목록과 같은 순서의 OnKey 이름 목록.
Private Const KEY_SHOW As String = "Enter Tab Esc Delete Backspace Home End PageUp PageDown Insert Up Down Left Right"
Private Const KEY_ONKEY As String = "ENTER TAB ESCAPE DELETE BACKSPACE HOME END PGUP PGDN INSERT UP DOWN LEFT RIGHT"
Private Const KEY_ALIAS As String = "ESC=ESCAPE DEL=DELETE BS=BACKSPACE BKSP=BACKSPACE RETURN=ENTER PAGEUP=PGUP PAGEDOWN=PGDN INS=INSERT"

Private mReg As Collection      ' 항목: onKey & vbTab & owner & vbTab & macro, 키: onKey

' 사람이 읽는 정규형("Ctrl+Shift+H"). 읽을 수 없으면 "".
Public Function CoKey_Normalize(ByVal keyText As String) As String
    Dim c As Boolean, s As Boolean, a As Boolean, k As String
    If Not ParseKey(keyText, c, s, a, k) Then Exit Function
    CoKey_Normalize = IIf(c, "Ctrl+", "") & IIf(s, "Shift+", "") & IIf(a, "Alt+", "") & ShowName(k)
End Function

' Application.OnKey 형식 정규형("^+H", "{F5}"). 읽을 수 없으면 "".
Public Function CoKey_ToOnKey(ByVal keyText As String) As String
    Dim c As Boolean, s As Boolean, a As Boolean, k As String
    If Not ParseKey(keyText, c, s, a, k) Then Exit Function
    CoKey_ToOnKey = IIf(c, "^", "") & IIf(s, "+", "") & IIf(a, "%", "") & IIf(Len(k) = 1 And IsPlainChar(k), k, "{" & k & "}")
End Function

' 위험 키면 True와 이유. 읽을 수 없는 키도 True(이유: 형식).
Public Function CoKey_IsDangerous(ByVal keyText As String, ByRef reason As String) As Boolean
    Dim c As Boolean, s As Boolean, a As Boolean, k As String, ok As String
    reason = ""
    If Not ParseKey(keyText, c, s, a, k) Then
        reason = "단축키 형식을 읽을 수 없습니다: " & keyText & " (예: Ctrl+Shift+H)"
        CoKey_IsDangerous = True: Exit Function
    End If
    ok = CoKey_ToOnKey(keyText)
    If InStr(1, " " & DANGER_KEYS & " ", " " & ok & " ", vbBinaryCompare) > 0 Then
        reason = CoKey_Normalize(keyText) & "은(는) 엑셀이 기본으로 쓰는 키라 등록하지 않습니다."
        CoKey_IsDangerous = True
    ElseIf Not c And Not a Then
        If Len(k) = 1 Or IsWordName(k) Then
            reason = CoKey_Normalize(keyText) & "은(는) 글을 치거나 칸을 옮길 때마다 걸리는 키라 등록하지 않습니다. Ctrl이나 Alt를 함께 쓰세요."
            CoKey_IsDangerous = True
        End If
    End If
End Function

' 단축키를 등록한다. 형식 오류·위험 키·다른 소유자가 쓰는 키면 등록하지 않고 False와 이유.
' 같은 소유자가 같은 키를 다시 등록하면 매크로만 바꾼다. macroName은 "'파일.xlam'!매크로" 꼴을 권한다.
Public Function CoKey_Register(ByVal keyText As String, ByVal macroName As String, ByVal ownerId As String, _
                               ByRef reason As String) As Boolean
    Dim ok As String, cur As String
    reason = ""
    If Len(Trim$(ownerId)) = 0 Or Len(Trim$(macroName)) = 0 Then reason = "추가기능 id와 매크로 이름이 필요합니다.": Exit Function
    If CoKey_IsDangerous(keyText, reason) Then Exit Function
    ok = CoKey_ToOnKey(keyText)
    cur = RegOwner(ok)
    If Len(cur) > 0 And StrComp(cur, ownerId, vbTextCompare) <> 0 Then
        reason = CoKey_Normalize(keyText) & "은(는) 이미 '" & cur & "'이(가) 쓰고 있어 등록하지 않았습니다."
        Exit Function
    End If
    On Error GoTo Fail
    Application.OnKey ok, macroName
    On Error GoTo 0
    If Len(cur) > 0 Then mReg.Remove ok
    RegAdd ok, ownerId, macroName
    CoKey_Register = True
    Exit Function
Fail:
    reason = "단축키를 등록하지 못했습니다(" & Err.Description & ")."
End Function

' 등록을 풀고 키를 엑셀 기본 동작으로 되돌린다. 등록한 소유자가 아니거나 등록되지 않은 키면 False.
Public Function CoKey_Unregister(ByVal keyText As String, ByVal ownerId As String) As Boolean
    Dim ok As String
    ok = CoKey_ToOnKey(keyText)
    If Len(ok) = 0 Then Exit Function
    If StrComp(RegOwner(ok), ownerId, vbTextCompare) <> 0 Or Len(ownerId) = 0 Then Exit Function
    On Error Resume Next
    Application.OnKey ok
    mReg.Remove ok
    On Error GoTo 0
    CoKey_Unregister = True
End Function

' 그 소유자가 등록한 키를 모두 푼다. 푼 개수.
Public Function CoKey_UnregisterAll(ByVal ownerId As String) As Long
    Dim i As Long, p() As String
    If mReg Is Nothing Or Len(ownerId) = 0 Then Exit Function
    For i = mReg.Count To 1 Step -1
        p = Split(mReg(i), vbTab)
        If StrComp(p(1), ownerId, vbTextCompare) = 0 Then
            On Error Resume Next
            Application.OnKey p(0)
            On Error GoTo 0
            mReg.Remove i
            CoKey_UnregisterAll = CoKey_UnregisterAll + 1
        End If
    Next i
End Function

' 이미 등록된 키면 True와 쓰는 소유자 id(owners). 등록은 하지 않는다. 읽을 수 없는 키는 False.
Public Function CoKey_Conflicts(ByVal keyText As String, ByRef owners As String) As Boolean
    owners = RegOwner(CoKey_ToOnKey(keyText))
    CoKey_Conflicts = (Len(owners) > 0)
End Function

'--- 내부 ------------------------------------------------------
Private Sub RegAdd(ByVal ok As String, ByVal ownerId As String, ByVal macroName As String)
    If mReg Is Nothing Then Set mReg = New Collection
    mReg.Add ok & vbTab & ownerId & vbTab & macroName, ok
End Sub

Private Function RegOwner(ByVal ok As String) As String
    Dim v As String
    If mReg Is Nothing Or Len(ok) = 0 Then Exit Function
    On Error Resume Next
    v = mReg(ok)
    On Error GoTo 0
    If Len(v) > 0 Then RegOwner = Split(v, vbTab)(1)
End Function

' 표기 둘 중 하나를 읽어 수정키와 키 이름(글자 하나는 대문자, 이름은 OnKey 이름: F5, ENTER ...)으로.
Private Function ParseKey(ByVal keyText As String, ByRef c As Boolean, ByRef s As Boolean, ByRef a As Boolean, _
                          ByRef k As String) As Boolean
    Dim t As String
    c = False: s = False: a = False: k = ""
    t = Trim$(keyText)
    If Len(t) = 0 Then Exit Function
    If ParseOnKey(t, c, s, a, k) Then ParseKey = True: Exit Function
    c = False: s = False: a = False: k = ""
    ParseKey = ParseShow(Replace(t, " ", ""), c, s, a, k)
End Function

Private Function ParseOnKey(ByVal t As String, ByRef c As Boolean, ByRef s As Boolean, ByRef a As Boolean, _
                            ByRef k As String) As Boolean
    Dim ch As String, rest As String
    Do While Len(t) > 0
        ch = Left$(t, 1)
        If ch = "^" Then
            If c Then Exit Function
            c = True
        ElseIf ch = "+" Then
            If s Then Exit Function
            s = True
        ElseIf ch = "%" Then
            If a Then Exit Function
            a = True
        Else
            Exit Do
        End If
        t = Mid$(t, 2)
    Loop
    rest = t
    If Len(rest) >= 3 And Left$(rest, 1) = "{" And Right$(rest, 1) = "}" Then
        k = CanonName(Mid$(rest, 2, Len(rest) - 2))
    ElseIf Len(rest) = 1 Then
        If IsPlainChar(rest) Then k = UCase$(rest)
    End If
    ParseOnKey = (Len(k) > 0)
End Function

Private Function ParseShow(ByVal t As String, ByRef c As Boolean, ByRef s As Boolean, ByRef a As Boolean, _
                           ByRef k As String) As Boolean
    Dim lastPart As String, idx As Long, parts() As String, i As Long
    If t = "+" Then
        lastPart = "+": t = ""
    ElseIf Right$(t, 2) = "++" Then
        lastPart = "+": t = Left$(t, Len(t) - 2)
    Else
        idx = InStrRev(t, "+")
        lastPart = Mid$(t, idx + 1)
        If idx > 0 Then t = Left$(t, idx - 1) Else t = ""
    End If
    If Len(t) > 0 Then
        parts = Split(t, "+")
        For i = 0 To UBound(parts)
            Select Case LCase$(parts(i))
                Case "ctrl", "control"
                    If c Then Exit Function
                    c = True
                Case "shift"
                    If s Then Exit Function
                    s = True
                Case "alt"
                    If a Then Exit Function
                    a = True
                Case ""
                Case Else
                    Exit Function
            End Select
        Next i
    End If
    If Len(lastPart) = 1 Then
        If IsPlainChar(lastPart) Or lastPart = "+" Then k = UCase$(lastPart)
    Else
        k = CanonName(lastPart)
    End If
    ParseShow = (Len(k) > 0)
End Function

' 이름 키를 OnKey 이름으로("pageup" -> "PGUP", "f5" -> "F5"). 모르면 "". 한 글자는 그 글자(대문자).
Private Function CanonName(ByVal raw As String) As String
    Dim up As String, al() As String, i As Long, n As Long
    If Len(raw) = 1 Then
        If raw <> " " And AscW(raw) >= 32 Then CanonName = UCase$(raw)
        Exit Function
    End If
    up = UCase$(raw)
    al = Split(KEY_ALIAS, " ")
    For i = 0 To UBound(al)
        If Split(al(i), "=")(0) = up Then up = Split(al(i), "=")(1): Exit For
    Next i
    If Left$(up, 1) = "F" And Len(up) <= 3 And IsNumeric(Mid$(up, 2)) Then
        n = CLng(Val(Mid$(up, 2)))
        If n >= 1 And n <= 15 And CStr(n) = Mid$(up, 2) Then CanonName = up
        Exit Function
    End If
    If IsWordName(up) Then CanonName = up
End Function

Private Function IsWordName(ByVal k As String) As Boolean
    IsWordName = (InStr(1, " " & KEY_ONKEY & " ", " " & k & " ", vbBinaryCompare) > 0)
End Function

Private Function ShowName(ByVal k As String) As String
    Dim shows() As String, oks() As String, i As Long
    ShowName = k
    shows = Split(KEY_SHOW, " "): oks = Split(KEY_ONKEY, " ")
    For i = 0 To UBound(oks)
        If oks(i) = k Then ShowName = shows(i): Exit Function
    Next i
End Function

' OnKey에서 뜻이 있는 ^ + % ~ { } ( ) 와 공백·조절 문자가 아닌 한 글자.
Private Function IsPlainChar(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    If AscW(ch) <= 32 Then Exit Function
    IsPlainChar = (InStr(1, "^+%~{}()", ch, vbBinaryCompare) = 0)
End Function
