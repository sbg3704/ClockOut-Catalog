Attribute VB_Name = "CoSetting"
'==============================================================
' 기능: 칼퇴 공용 라이브러리 - 설정 읽기를 한 줄로(형 맞추기, 범위 제한, 선택지 확인).
'       표준 읽기 함수 ClockOut_GetSetting(ClockOutSettings.bas)을 감싼다. 새로 읽지 않는다.
'       image-fit의 LoadOptions에서 매번 쓰던 "읽고 -> 형 바꾸고 -> min/max로 한 번 더 제한"을 모았다.
' 사용법(값은 .clockout.json의 default·min·max·options와 같게 적는다):
'   mm  = CoSetting_Number(ADDIN_ID, "margin", 0.8, 0, 20)
'   pt  = CoSetting_MmAsPt(ADDIN_ID, "margin", 0.8, 0, 20)
'   n   = CLng(CoSetting_Number(ADDIN_ID, "count", 3, 1, 99))
'   ok  = CoSetting_Bool(ADDIN_ID, "keepRatio", True)
'   key = CoSetting_Text(ADDIN_ID, "keyInsert", "^+A")
'   al  = CoSetting_Choice(ADDIN_ID, "align", "가운데", "가운데", "왼쪽 위")   ' 목록에 없는 값이면 기본값
' 의존: ClockOutSettings.bas(같은 폴더), CoNum.bas. 못 읽으면 항상 기본값(오류를 던지지 않음).
'==============================================================
Option Explicit
Option Private Module

' 숫자 설정. lo·hi를 주면 그 안으로 제한한다.
Public Function CoSetting_Number(ByVal addinId As String, ByVal key As String, ByVal defaultValue As Double, _
                                 Optional ByVal lo As Variant, Optional ByVal hi As Variant) As Double
    Dim v As Double
    On Error GoTo Fallback
    v = CDbl(ClockOut_GetSetting(addinId, key, defaultValue))
    If Not IsMissing(lo) Then
        If v < CDbl(lo) Then v = CDbl(lo)
    End If
    If Not IsMissing(hi) Then
        If v > CDbl(hi) Then v = CDbl(hi)
    End If
    CoSetting_Number = v
    Exit Function
Fallback:
    CoSetting_Number = defaultValue
End Function

' mm로 선언한 설정을 pt로 읽는다(도형 크기·위치는 pt).
Public Function CoSetting_MmAsPt(ByVal addinId As String, ByVal key As String, ByVal defaultMm As Double, _
                                 Optional ByVal loMm As Variant, Optional ByVal hiMm As Variant) As Double
    CoSetting_MmAsPt = CoNum_MmToPt(CoSetting_Number(addinId, key, defaultMm, loMm, hiMm))
End Function

Public Function CoSetting_Bool(ByVal addinId As String, ByVal key As String, ByVal defaultValue As Boolean) As Boolean
    On Error GoTo Fallback
    CoSetting_Bool = CBool(ClockOut_GetSetting(addinId, key, defaultValue))
    Exit Function
Fallback:
    CoSetting_Bool = defaultValue
End Function

' 글자 설정. 기본은 앞뒤 공백을 지운다.
Public Function CoSetting_Text(ByVal addinId As String, ByVal key As String, ByVal defaultValue As String, _
                               Optional ByVal trimIt As Boolean = True) As String
    Dim s As String
    On Error GoTo Fallback
    s = CStr(ClockOut_GetSetting(addinId, key, defaultValue))
    If trimIt Then s = Trim$(s)
    CoSetting_Text = s
    Exit Function
Fallback:
    CoSetting_Text = defaultValue
End Function

' 선택지 설정. 읽은 값이 options 중 하나와 정확히 같을 때만 돌려주고, 아니면 기본값.
' options를 하나도 안 주면 읽은 값을 그대로 돌려준다.
Public Function CoSetting_Choice(ByVal addinId As String, ByVal key As String, ByVal defaultValue As String, _
                                 ParamArray options() As Variant) As String
    Dim s As String, i As Long
    On Error GoTo Fallback
    s = CStr(ClockOut_GetSetting(addinId, key, defaultValue))
    If UBound(options) < LBound(options) Then CoSetting_Choice = s: Exit Function
    For i = LBound(options) To UBound(options)
        If s = CStr(options(i)) Then CoSetting_Choice = s: Exit Function
    Next i
Fallback:
    CoSetting_Choice = defaultValue
End Function
