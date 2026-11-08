Attribute VB_Name = "CoShape"
'==============================================================
' 기능: 칼퇴 공용 라이브러리 - 사진·도형 다루기(선택 얻기, 안전하게 넣기·지우기, 칸에 맞추기).
' 사용법:
'   Dim sr As ShapeRange
'   If Not CoShape_GetSelected(sr) Then CoGuard_Notice "도형을 먼저 선택하세요.": Exit Sub
'   Set shp = CoShape_AddPicture(ws, 경로, why)          ' 실패하면 Nothing, why에 이유
'   ok = CoShape_FitToRange(shp, 칸, 여백pt, 비율유지, 가운데)
'   CoShape_Delete shp                                    ' 오류 처리부에서 반쯤 놓인 개체 정리
' 의존: CoFitMath.bas(같은 폴더, 맞춤 계산). 추가기능 폴더의 파일에는 기대지 않는다.
'==============================================================
Option Explicit
Option Private Module

' 선택된 도형들을 얻는다. 셀이 선택돼 있거나 아무것도 없으면 False.
Public Function CoShape_GetSelected(ByRef sr As ShapeRange) As Boolean
    Set sr = Nothing
    On Error Resume Next
    Set sr = Selection.ShapeRange                     ' 도형이 아닌 선택이면 오류
    Err.Clear
    On Error GoTo 0
    If sr Is Nothing Then Exit Function
    CoShape_GetSelected = (sr.Count > 0)
End Function

' 사진 파일을 원본 크기로 넣는다. 실패하면 Nothing을 돌려주고 reason에 이유를 적는다.
Public Function CoShape_AddPicture(ByVal ws As Worksheet, ByVal filePath As String, ByRef reason As String) As Shape
    Dim errNo As Long
    reason = ""
    On Error Resume Next
    Set CoShape_AddPicture = ws.Shapes.AddPicture(filePath, msoFalse, msoTrue, 0, 0, -1, -1)   ' -1 = 원본 크기
    errNo = Err.Number
    On Error GoTo 0
    If errNo <> 0 Then
        Set CoShape_AddPicture = Nothing
        reason = "사진을 불러오지 못했습니다." & vbCrLf & _
                 "(파일이 없거나 손상됐거나 Excel이 지원하지 않는 형식일 수 있습니다)" & vbCrLf & filePath
    End If
End Function

' beforeCount(작업 전 ws.Shapes.Count) 뒤에 새로 생긴 맨 위 도형. 없으면 Nothing.
' 예: n = ws.Shapes.Count : ws.Pictures.Paste : Set shp = CoShape_LastAdded(ws, n)
Public Function CoShape_LastAdded(ByVal ws As Worksheet, ByVal beforeCount As Long) As Shape
    If ws.Shapes.Count > beforeCount Then Set CoShape_LastAdded = ws.Shapes(ws.Shapes.Count)
End Function

' 도형을 지운다. Nothing이거나 이미 지워졌어도 오류가 나지 않는다.
Public Sub CoShape_Delete(ByRef shp As Shape)
    On Error Resume Next
    If Not shp Is Nothing Then shp.Delete
    Set shp = Nothing
End Sub

' 기울지 않은(0도·180도) 도형이면 True. 90도 등은 가로·세로가 뒤섞여 칸 맞춤에서 뺀다.
Public Function CoShape_IsUpright(ByVal shp As Shape) As Boolean
    Dim r As Double
    r = shp.Rotation
    CoShape_IsUpright = (r = 0 Or r = 180 Or r = 360)
End Function

' 사진이면 원본 크기로 되돌려 원래 비율을 되찾는다. 사진이 아니면 아무것도 안 한다.
Public Sub CoShape_RestoreOriginalSize(ByVal shp As Shape)
    On Error Resume Next
    If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
        shp.ScaleHeight 1, msoTrue
        shp.ScaleWidth 1, msoTrue
    End If
End Sub

' 도형을 칸(area)에 맞춘다. 맞출 수 없으면(기운 도형, 숨긴 칸) 아무것도 안 바꾸고 False.
'   marginPt: 사방 여백(pt). 칸이 작으면 짧은 변의 1/4까지만 적용(CoFitMath 규칙).
'   placement: xlMoveAndSize·xlMove·xlFreeFloating. 0이면 바꾸지 않는다.
Public Function CoShape_FitToRange(ByVal shp As Shape, ByVal area As Range, ByVal marginPt As Double, _
                                   Optional ByVal keepRatio As Boolean = True, _
                                   Optional ByVal alignCenter As Boolean = True, _
                                   Optional ByVal placement As Long = 0) As Boolean
    Dim box As CoFitBox
    If shp Is Nothing Or area Is Nothing Then Exit Function
    If Not CoShape_IsUpright(shp) Then Exit Function

    box = CoFitMath_CalcBox(area.Left, area.Top, area.Width, area.Height, marginPt, _
                            shp.Width, shp.Height, keepRatio, alignCenter)
    If Not box.Ok Then Exit Function

    shp.LockAspectRatio = msoFalse                    ' 켜져 있으면 너비만 바꿔도 높이가 따라 바뀐다
    shp.Width = box.Width
    shp.Height = box.Height
    shp.Left = box.Left
    shp.Top = box.Top
    If keepRatio Then shp.LockAspectRatio = msoTrue
    If placement <> 0 Then shp.Placement = placement
    CoShape_FitToRange = True
End Function

' area 안에 왼쪽 위가 놓인 도형들(Collection of Shape). 같은 시트의 범위여야 한다.
Public Function CoShape_InRange(ByVal area As Range) As Collection
    Dim col As New Collection, shp As Shape, hit As Range
    For Each shp In area.Worksheet.Shapes
        Set hit = Nothing
        On Error Resume Next
        Set hit = Application.Intersect(shp.TopLeftCell, area)   ' 일부 도형(메모 등)은 TopLeftCell 오류
        On Error GoTo 0
        If Not hit Is Nothing Then col.Add shp
    Next shp
    Set CoShape_InRange = col
End Function
