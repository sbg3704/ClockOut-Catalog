Attribute VB_Name = "CoKeyTest"
'==============================================================
' 기능: CoKey(단축키 등록부) 시험. 정규화·위험 키 판정은 순수 계산이고, 등록·충돌·해제는
'       잘 안 쓰는 Ctrl+Shift+Alt+F15로 Excel 안에서 실제 OnKey를 걸었다가 푼다(끝나면 남는 것 없음).
' 사용법: 직접 실행 창에서 Test_CoKey  (빌드: tools\build_addin.ps1 -Addin _common -RunTests)
' 의존: CoKey.bas
'==============================================================
Option Explicit
Option Private Module

Public Sub Test_CoKey()
    Dim f As Long, why As String, who As String
    Const K As String = "Ctrl+Shift+Alt+F15"
    ' 정규화: 대소문자·순서·공백 무시, 두 표기 모두
    f = f + Chk("표시형", CoKey_Normalize("ctrl + shift + c") = "Ctrl+Shift+C")
    f = f + Chk("순서", CoKey_Normalize("shift+ctrl+h") = "Ctrl+Shift+H")
    f = f + Chk("OnKey 입력", CoKey_Normalize("^+h") = "Ctrl+Shift+H")
    f = f + Chk("OnKey 출력", CoKey_ToOnKey("Ctrl+Shift+H") = "^+H")
    f = f + Chk("이름 키", CoKey_ToOnKey("alt+pageup") = "%{PGUP}")
    f = f + Chk("별칭", CoKey_Normalize("ctrl+esc") = "Ctrl+Esc")
    f = f + Chk("F키", CoKey_ToOnKey("{f5}") = "{F5}")
    f = f + Chk("더하기 키", CoKey_ToOnKey("Ctrl++") = "^{+}")
    f = f + Chk("빈 입력", CoKey_Normalize("") = "")
    f = f + Chk("수정키만", CoKey_Normalize("Ctrl+Shift") = "")
    f = f + Chk("수정키 두 번", CoKey_Normalize("Ctrl+Ctrl+H") = "")
    f = f + Chk("모르는 이름", CoKey_Normalize("Ctrl+Banana") = "")
    f = f + Chk("F16 없음", CoKey_Normalize("F16") = "")
    ' 위험 키
    f = f + Chk("Ctrl+C 위험", CoKey_IsDangerous("Ctrl+C", why) And Len(why) > 0)
    f = f + Chk("Ctrl+Shift+Z 위험", CoKey_IsDangerous("^+z", why))
    f = f + Chk("Ctrl+Alt+V 위험", CoKey_IsDangerous("Ctrl+Alt+V", why))
    f = f + Chk("F5 단독 위험", CoKey_IsDangerous("F5", why))
    f = f + Chk("Alt+F4 위험", CoKey_IsDangerous("Alt+F4", why))
    f = f + Chk("글자 하나 위험", CoKey_IsDangerous("h", why))
    f = f + Chk("Shift+글자 위험", CoKey_IsDangerous("Shift+H", why))
    f = f + Chk("Enter 단독 위험", CoKey_IsDangerous("Enter", why))
    f = f + Chk("형식 오류 위험", CoKey_IsDangerous("Ctrl+", why))
    f = f + Chk("Ctrl+Shift+H 안전", Not CoKey_IsDangerous("Ctrl+Shift+H", why) And why = "")
    f = f + Chk("Shift+F5 안전", Not CoKey_IsDangerous("Shift+F5", why))
    ' 등록부(Excel 안에서 실제 OnKey)
    CoKey_UnregisterAll "test-a": CoKey_UnregisterAll "test-b"
    f = f + Chk("등록", CoKey_Register(K, "CoKeyTest_Nothing", "test-a", why))
    f = f + Chk("같은 소유자 다시", CoKey_Register("^+%{F15}", "CoKeyTest_Nothing", "test-a", why))
    f = f + Chk("다른 소유자 거부", Not CoKey_Register(K, "CoKeyTest_Nothing", "test-b", why) And InStr(why, "test-a") > 0)
    f = f + Chk("충돌 확인", CoKey_Conflicts("alt+shift+ctrl+f15", who) And who = "test-a")
    f = f + Chk("위험 키 등록 거부", Not CoKey_Register("Ctrl+V", "CoKeyTest_Nothing", "test-a", why) And Len(why) > 0)
    f = f + Chk("남이 해제 못 함", Not CoKey_Unregister(K, "test-b"))
    f = f + Chk("해제", CoKey_Unregister(K, "test-a"))
    f = f + Chk("해제 뒤 비어 있음", Not CoKey_Conflicts(K, who) And who = "")
    f = f + Chk("모두 해제", CoKey_Register(K, "CoKeyTest_Nothing", "test-b", why) And CoKey_UnregisterAll("test-b") = 1)
    If f = 0 Then
        Debug.Print "CoKey 시험: 모두 통과"
        MsgBox "CoKey 시험: 모두 통과", vbInformation, "칼퇴"
    Else
        MsgBox "CoKey 시험: " & f & "개 실패 (직접 실행 창 확인)", vbExclamation, "칼퇴"
    End If
End Sub

Private Function Chk(ByVal testName As String, ByVal ok As Boolean) As Long
    If Not ok Then Debug.Print "[실패] " & testName: Chk = 1
End Function
