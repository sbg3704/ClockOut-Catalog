VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoVendorRegex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================
' 기능: stdRegex를 만드는 팩토리. stdRegex.Create의 기본값은 실험용 기계어 실행 모드(useNative=True)라
'       Excel이 꺼질 수 있고 백신에 걸릴 수도 있어서, 여기서는 항상 순수 VBA 모드(useNative=False)로 만든다.
' 사용법:
'   Set rx = v.CoVendor_Regex().Create("(\d+)-(\d+)")       ' 플래그: "i" 대소문자 무시, "m" 여러 줄, "s" 점이 줄바꿈 포함, "g" 전체
'   rx.Test(s) / rx.Match(s) / rx.MatchAll(s) / rx.Replace(s, "$2-$1")
' 의존: stdRegex.cls
'==============================================================
Option Explicit

' 정규식 개체(stdRegex)를 만든다. stepsLimit은 한 번 맞춰 볼 때 최대 단계 수(넘으면 원본처럼 오류).
Public Function Create(ByVal pattern As String, Optional ByVal flags As String = "", _
                       Optional ByVal stepsLimit As Long = 10000) As Object
    Set Create = stdRegex.Create(pattern, flags, stepsLimit, False)
End Function
