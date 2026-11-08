Attribute VB_Name = "ClockOutSettings"
'==============================================================
' 기능: 칼퇴 매니저 설정창에서 저장한 값을 읽는 표준 읽기 모듈 (초안 v0).
'       %APPDATA%\ClockOut\settings\<추가기능 id>.json 파일(평평한 JSON)에서 값을 읽는다.
' 사용법: v = ClockOut_GetSetting("image-fit", "margin", 0.8)
'         - 마지막 인자(기본값)의 형이 곧 돌려받을 형이다. Double/Long -> 숫자, Boolean -> 참거짓, String -> 글자.
'         - 파일이 없거나, 깨졌거나, 키가 없거나, 형이 안 맞으면 항상 기본값을 돌려준다(오프라인·안전 원칙).
'         - 값을 못 읽어도 오류를 밖으로 던지지 않는다.
' 의존: Windows(ADODB.Stream을 후기 바인딩으로 사용, 참조 추가 불필요). 없으면 기본값으로 동작한다.
' 주의: docs/03 초안은 ClockOut_GetSetting("margin", 1.0)처럼 id 없이 썼으나, 읽기 모듈을 여러
'       추가기능이 함께 쓰므로 id를 첫 인자로 받는다. 확정은 구현하면서 한다(docs/03 참고).
'==============================================================
Option Explicit

Private Const KIND_STRING As Long = 1
Private Const KIND_NUMBER As Long = 2
Private Const KIND_BOOL As Long = 3
Private Const KIND_OTHER As Long = 4

Private mCacheId As String
Private mCacheStamp As Date
Private mCacheText As String
Private mCacheOk As Boolean

Public Function ClockOut_GetSetting(ByVal addinId As String, ByVal key As String, ByVal defaultValue As Variant) As Variant
    Dim txt As String, kind As Long, raw As String, result As Variant
    ClockOut_GetSetting = defaultValue
    On Error GoTo Fallback
    txt = LoadSettingsText(addinId)
    If Len(txt) = 0 Then Exit Function
    If Not FindTopLevelValue(txt, key, kind, raw) Then Exit Function
    If Coerce(defaultValue, kind, raw, result) Then ClockOut_GetSetting = result
    Exit Function
Fallback:
    ClockOut_GetSetting = defaultValue
End Function

'--- 파일 읽기 -------------------------------------------------
Private Function SettingsPath(ByVal addinId As String) As String
    Dim root As String
    root = Environ$("APPDATA")
    If Len(root) = 0 Then Exit Function
    SettingsPath = root & "\ClockOut\settings\" & SafeName(addinId) & ".json"
End Function

' 매니저(UserSettingsStore.SafeName)와 같은 규칙: 영문·숫자·한글·-·_ 외에는 _로 바꾼다.
Private Function SafeName(ByVal id As String) As String
    Dim i As Long, c As String, code As Long, outText As String
    For i = 1 To Len(id)
        c = Mid$(id, i, 1)
        code = AscW(c)
        If code < 0 Then code = code + 65536
        If (c >= "0" And c <= "9") Or (c >= "a" And c <= "z") Or (c >= "A" And c <= "Z") _
           Or c = "-" Or c = "_" Or (code >= &HAC00 And code <= &HD7A3) Then
            outText = outText & c
        Else
            outText = outText & "_"
        End If
    Next i
    SafeName = outText
End Function

Private Function LoadSettingsText(ByVal addinId As String) As String
    Dim p As String, stamp As Date, st As Object, txt As String
    p = SettingsPath(addinId)
    If Len(p) = 0 Then Exit Function
    On Error GoTo Fail
    stamp = FileDateTime(p)                       ' 파일이 없으면 오류 -> Fail
    If mCacheOk And mCacheId = addinId And mCacheStamp = stamp Then
        LoadSettingsText = mCacheText
        Exit Function
    End If
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                                   ' 글자
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile p
    txt = st.ReadText
    st.Close
    If Len(txt) > 0 Then
        If AscW(Left$(txt, 1)) = &HFEFF Then txt = Mid$(txt, 2)   ' BOM 제거
    End If
    mCacheId = addinId: mCacheStamp = stamp: mCacheText = txt: mCacheOk = True
    LoadSettingsText = txt
    Exit Function
Fail:
    mCacheOk = False
    LoadSettingsText = ""
End Function

'--- 평평한 JSON에서 한 키의 값 찾기 ---------------------------------
Private Function FindTopLevelValue(ByVal s As String, ByVal wantKey As String, _
                                   ByRef kind As Long, ByRef raw As String) As Boolean
    Dim i As Long, n As Long, ch As String, k As String, v As String, vk As Long
    n = Len(s): i = 1
    SkipWs s, i
    If i > n Then Exit Function
    If Mid$(s, i, 1) <> "{" Then Exit Function
    i = i + 1
    Do
        SkipWs s, i
        If i > n Then Exit Function
        ch = Mid$(s, i, 1)
        If ch = "}" Then Exit Function
        If ch = "," Then
            i = i + 1
        ElseIf ch = """" Then
            k = ReadJsonString(s, i)
            SkipWs s, i
            If i > n Then Exit Function
            If Mid$(s, i, 1) <> ":" Then Exit Function
            i = i + 1
            SkipWs s, i
            If i > n Then Exit Function
            If Mid$(s, i, 1) = """" Then
                v = ReadJsonString(s, i): vk = KIND_STRING
            Else
                v = ReadBareValue(s, i, vk)
            End If
            If k = wantKey Then
                kind = vk: raw = v
                FindTopLevelValue = True
                Exit Function
            End If
        Else
            Exit Function                          ' 형식이 깨짐 -> 기본값
        End If
    Loop
End Function

Private Sub SkipWs(ByVal s As String, ByRef i As Long)
    Dim c As String
    Do While i <= Len(s)
        c = Mid$(s, i, 1)
        If c = " " Or c = vbTab Or c = vbCr Or c = vbLf Then i = i + 1 Else Exit Do
    Loop
End Sub

' i 위치가 여는 따옴표. 읽은 뒤 i는 닫는 따옴표 다음.
Private Function ReadJsonString(ByVal s As String, ByRef i As Long) As String
    Dim outText As String, c As String, n As Long
    n = Len(s)
    i = i + 1
    Do While i <= n
        c = Mid$(s, i, 1)
        If c = """" Then
            i = i + 1
            ReadJsonString = outText
            Exit Function
        ElseIf c = "\" Then
            i = i + 1
            If i > n Then Exit Do
            c = Mid$(s, i, 1)
            Select Case c
                Case "n": outText = outText & vbLf
                Case "r": outText = outText & vbCr
                Case "t": outText = outText & vbTab
                Case "b": outText = outText & Chr$(8)
                Case "f": outText = outText & Chr$(12)
                Case "u"
                    If i + 4 > n Then Exit Do
                    outText = outText & ChrW(CLng("&H" & Mid$(s, i + 1, 4)))
                    i = i + 4
                Case Else: outText = outText & c           ' \" \\ \/ 등
            End Select
            i = i + 1
        Else
            outText = outText & c
            i = i + 1
        End If
    Loop
    Err.Raise 5, , "닫히지 않은 JSON 글자"         ' GetSetting의 오류 처리가 기본값으로 돌린다
End Function

' 숫자·true·false·null, 그리고 건너뛸 중첩 값([...], {...}).
Private Function ReadBareValue(ByVal s As String, ByRef i As Long, ByRef kind As Long) As String
    Dim c As String, start As Long, depth As Long, inQuote As Boolean, tok As String
    c = Mid$(s, i, 1)
    If c = "[" Or c = "{" Then
        Do While i <= Len(s)
            c = Mid$(s, i, 1)
            If inQuote Then
                If c = "\" Then
                    i = i + 1
                ElseIf c = """" Then
                    inQuote = False
                End If
            ElseIf c = """" Then
                inQuote = True
            ElseIf c = "[" Or c = "{" Then
                depth = depth + 1
            ElseIf c = "]" Or c = "}" Then
                depth = depth - 1
                If depth = 0 Then i = i + 1: Exit Do
            End If
            i = i + 1
        Loop
        kind = KIND_OTHER
        Exit Function
    End If
    start = i
    Do While i <= Len(s)
        c = Mid$(s, i, 1)
        If c = "," Or c = "}" Or c = " " Or c = vbTab Or c = vbCr Or c = vbLf Then Exit Do
        i = i + 1
    Loop
    tok = Mid$(s, start, i - start)
    Select Case tok
        Case "true", "false": kind = KIND_BOOL
        Case "null": kind = KIND_OTHER
        Case Else: kind = KIND_NUMBER
    End Select
    ReadBareValue = tok
End Function

'--- 기본값의 형에 맞게 바꾸기 -----------------------------------------
Private Function Coerce(ByVal defaultValue As Variant, ByVal kind As Long, _
                        ByVal raw As String, ByRef result As Variant) As Boolean
    Select Case VarType(defaultValue)
        Case vbBoolean
            If kind = KIND_BOOL Or kind = KIND_STRING Then
                If LCase$(raw) = "true" Then
                    result = True: Coerce = True
                ElseIf LCase$(raw) = "false" Then
                    result = False: Coerce = True
                End If
            End If
        Case vbDouble, vbSingle, vbInteger, vbLong, vbCurrency
            If kind = KIND_NUMBER Then
                result = Val(raw)                  ' Val은 지역 설정과 무관하게 '.'을 소수점으로 읽는다
                Coerce = True
            End If
        Case vbString
            If kind = KIND_STRING Then
                result = raw: Coerce = True
            End If
    End Select
End Function
