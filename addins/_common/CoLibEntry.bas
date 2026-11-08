Attribute VB_Name = "CoLibEntry"
'==============================================================
' 기능: 칼퇴 공용 라이브러리(clockout-lib.xlam)의 바깥 입구. 이 모듈만 Option Private Module이 없어
'       다른 통합 문서(추가기능)가 Application.Run으로 부를 수 있다.
' 사용법(추가기능 쪽):
'   ver = Application.Run("'clockout-lib.xlam'!CoLib_Version")     ' "0.1.000" (clockout-lib.clockout.json의 version과 같음)
'   Set lib = Application.Run("'clockout-lib.xlam'!CoLib_Api")     ' CoLibApi 개체(하나만 만들어 계속 같은 것을 돌려줌)
'   lib.CoText_CleanInvisible(s) 처럼 늦은 바인딩(Dim lib As Object)으로 부른다. 메서드 목록은 CoLibApi.cls.
'   시험: Test_CoLibApi
' 주의: 판을 올릴 때 LIB_VERSION과 clockout-lib.clockout.json의 version을 함께 고친다.
' 의존: CoLibApi.cls(같은 폴더)와 그것이 감싸는 모든 Co 모듈
'==============================================================
Option Explicit

Private Const LIB_VERSION As String = "0.1.000"
Private mApi As CoLibApi

Public Function CoLib_Version() As String
    CoLib_Version = LIB_VERSION
End Function

Public Function CoLib_Api() As Object
    If mApi Is Nothing Then Set mApi = New CoLibApi
    Set CoLib_Api = mApi
End Function

'--- 시험 (Excel 안에서 직접 실행: 보기 > 직접 실행 창) --------------------
' 개체를 Application.Run으로 얻어 늦은 바인딩으로 부르고, ByRef 인자(Double·Date·String)가 돌아오는지 본다.
Public Sub Test_CoLibApi()
    Dim f As Long, lib As Object, x As Double, dt As Date, txt As String, v As Variant, tok As Long
    Set lib = Application.Run("'" & ThisWorkbook.Name & "'!CoLib_Api")
    f = f + Chk("같은 개체", lib Is CoLib_Api())
    f = f + Chk("판", Application.Run("'" & ThisWorkbook.Name & "'!CoLib_Version") = "0.1.000")
    f = f + Chk("CoNum_Clamp", lib.CoNum_Clamp(25, 0, 20) = 20)
    f = f + Chk("CoText_ColLetter", lib.CoText_ColLetter(28) = "AB")
    f = f + Chk("CoText_CleanInvisible", lib.CoText_CleanInvisible("  a   b ") = "a b")
    f = f + Chk("CoNum_ToKoreanWon", lib.CoNum_ToKoreanWon(1234) = "일금 일천이백삼십사원정")
    f = f + Chk("ByRef Double", lib.CoNum_TryParse("1.5", x) And x = 1.5)
    f = f + Chk("ByRef Date", lib.CoText_ParseDate("2026.11.08", dt) And dt = DateSerial(2026, 11, 8))
    txt = "x"
    f = f + Chk("ByRef String", Not lib.CoFile_ReadUtf8("Z:\없는\폴더\없는파일.txt", txt) And txt = "")
    v = lib.CoFitMath_CalcBox(0, 0, 100, 50, 0, 100, 100, True, True)
    f = f + Chk("CalcBox 배열", IsArray(v))
    If IsArray(v) Then f = f + Chk("CalcBox 값", Abs(v(0) - 25) < 0.000001 And Abs(v(2) - 50) < 0.000001)
    f = f + Chk("CalcBox 못 맞춤", IsEmpty(lib.CoFitMath_CalcBox(0, 0, 0, 50, 0, 100, 100, True, True)))
    tok = lib.CoGuard_Begin(False)
    lib.CoGuard_End tok
    lib.CoGuard_End tok                                    ' 두 번 불러도 안전
    f = f + Chk("CoGuard 토큰", tok >= 1)
    f = f + Chk("상수", lib.FilterAll = CO_FILTER_ALL And lib.MsgTitle = CO_MSG_TITLE)
    f = f + Chk("CoSetting_Choice 기본값", lib.CoSetting_Choice("없는-추가기능-id", "align", "가운데", "가운데", "왼쪽 위") = "가운데")
    If f = 0 Then
        Debug.Print "CoLibApi 시험: 모두 통과"
        MsgBox "CoLibApi 시험: 모두 통과", vbInformation, "칼퇴"
    Else
        MsgBox "CoLibApi 시험: " & f & "개 실패 (직접 실행 창 확인)", vbExclamation, "칼퇴"
    End If
End Sub

Private Function Chk(ByVal testName As String, ByVal ok As Boolean) As Long
    If Not ok Then Debug.Print "실패: " & testName: Chk = 1
End Function
