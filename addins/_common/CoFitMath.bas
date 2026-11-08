Attribute VB_Name = "CoFitMath"
'==============================================================
' 기능: 칼퇴 공용 라이브러리 - 사진·도형을 칸(셀 영역) 안에 맞출 때의 위치·크기 계산(순수 계산).
'       화면·Office 개체를 쓰지 않아 Excel 없이도 같은 입력이면 같은 답이 나온다.
'       image-fit의 ImageFitMath.CalcFitBox와 같은 규칙이다. 공용은 이쪽이 원본이고,
'       image-fit은 PR#10이 합쳐진 뒤 이 모듈을 쓰도록 바꾼다(후속).
' 사용법: box = CoFitMath_CalcBox(칸왼쪽, 칸위, 칸너비, 칸높이, 여백, 도형너비, 도형높이, 비율유지, 가운데정렬)
'         box.Ok가 False면 맞출 수 없는 칸(숨김·너무 작음)이다. 단위는 모두 포인트(pt).
'         여백은 사방 같고, 칸이 작으면 짧은 변의 1/4까지만 적용한다.
'         시험: Test_CoFitMath (직접 실행 창에 결과가 나온다).
' 의존: 없음
'==============================================================
Option Explicit
Option Private Module

Public Type CoFitBox
    Ok As Boolean
    Left As Double
    Top As Double
    Width As Double
    Height As Double
End Type

Public Function CoFitMath_CalcBox(ByVal areaLeft As Double, ByVal areaTop As Double, _
                                  ByVal areaW As Double, ByVal areaH As Double, _
                                  ByVal marginPt As Double, _
                                  ByVal shapeW As Double, ByVal shapeH As Double, _
                                  ByVal keepRatio As Boolean, ByVal alignCenter As Boolean) As CoFitBox
    Dim m As Double, shortSide As Double, boxW As Double, boxH As Double
    Dim fitW As Double, fitH As Double, fitScale As Double, sx As Double, sy As Double
    Dim r As CoFitBox

    If areaW <= 0 Or areaH <= 0 Then                  ' 숨긴 행·열이거나 크기 0인 칸
        CoFitMath_CalcBox = r
        Exit Function
    End If

    m = marginPt
    If m < 0 Then m = 0
    shortSide = areaW
    If areaH < shortSide Then shortSide = areaH
    If m > shortSide / 4 Then m = shortSide / 4       ' 작은 칸에서 여백 때문에 사진이 사라지지 않게

    boxW = areaW - 2 * m
    boxH = areaH - 2 * m
    If boxW <= 0 Or boxH <= 0 Then
        CoFitMath_CalcBox = r
        Exit Function
    End If

    If keepRatio And shapeW > 0 And shapeH > 0 Then
        sx = boxW / shapeW
        sy = boxH / shapeH
        If sx < sy Then fitScale = sx Else fitScale = sy
        fitW = shapeW * fitScale
        fitH = shapeH * fitScale
    Else                                              ' 비율을 안 지키거나 크기를 모를 때는 칸에 꽉 채움
        fitW = boxW
        fitH = boxH
    End If

    r.Ok = True
    r.Width = fitW
    r.Height = fitH
    If alignCenter Then
        r.Left = areaLeft + m + (boxW - fitW) / 2
        r.Top = areaTop + m + (boxH - fitH) / 2
    Else
        r.Left = areaLeft + m
        r.Top = areaTop + m
    End If
    CoFitMath_CalcBox = r
End Function

'--- 시험 (Excel 안에서 직접 실행: 보기 > 직접 실행 창) --------------------
Public Sub Test_CoFitMath()
    Dim f As Long, b As CoFitBox

    b = CoFitMath_CalcBox(0, 0, 100, 50, 0, 200, 100, True, True)      ' 비율이 칸과 같으면 꽉 참
    f = f + Chk("같은 비율", b.Ok And Near(b.Width, 100) And Near(b.Height, 50) And Near(b.Left, 0) And Near(b.Top, 0))

    b = CoFitMath_CalcBox(0, 0, 100, 50, 0, 100, 100, True, True)      ' 정사각 -> 가운데 50x50
    f = f + Chk("가운데 정렬", b.Ok And Near(b.Width, 50) And Near(b.Height, 50) And Near(b.Left, 25) And Near(b.Top, 0))

    b = CoFitMath_CalcBox(0, 0, 100, 50, 0, 100, 100, True, False)     ' 왼쪽 위 정렬
    f = f + Chk("왼쪽 위 정렬", b.Ok And Near(b.Width, 50) And Near(b.Left, 0) And Near(b.Top, 0))

    b = CoFitMath_CalcBox(10, 20, 100, 50, 2, 300, 10, False, True)    ' 비율 무시 -> 여백만 뺀 꽉 채움
    f = f + Chk("늘려 채우기", b.Ok And Near(b.Width, 96) And Near(b.Height, 46) And Near(b.Left, 12) And Near(b.Top, 22))

    b = CoFitMath_CalcBox(0, 0, 10, 8, 10, 5, 5, False, True)          ' 여백이 너무 큼 -> 짧은 변의 1/4(2pt)
    f = f + Chk("여백 제한", b.Ok And Near(b.Width, 6) And Near(b.Height, 4))

    b = CoFitMath_CalcBox(0, 0, 0, 50, 0, 100, 100, True, True)        ' 너비 0(숨김)
    f = f + Chk("숨김 칸", Not b.Ok)

    b = CoFitMath_CalcBox(0, 0, 100, 50, 0, 0, 0, True, True)          ' 크기를 모를 때
    f = f + Chk("도형 크기 0", b.Ok And Near(b.Width, 100) And Near(b.Height, 50))

    b = CoFitMath_CalcBox(0, 0, 100, 50, -5, 200, 100, True, True)     ' 음수 여백은 0으로
    f = f + Chk("음수 여백", b.Ok And Near(b.Width, 100))

    If f = 0 Then
        Debug.Print "CoFitMath 시험: 모두 통과"
        MsgBox "CoFitMath 시험: 모두 통과", vbInformation, "칼퇴"
    Else
        MsgBox "CoFitMath 시험: " & f & "개 실패 (직접 실행 창 확인)", vbExclamation, "칼퇴"
    End If
End Sub

Private Function Near(ByVal a As Double, ByVal b As Double) As Boolean
    Near = (Abs(a - b) < 0.000001)
End Function

Private Function Chk(ByVal testName As String, ByVal ok As Boolean) As Long
    If Not ok Then Debug.Print "실패: " & testName: Chk = 1
End Function
