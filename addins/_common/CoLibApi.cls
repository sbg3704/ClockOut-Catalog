VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoLibApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==============================================================
' 기능: 칼퇴 공용 라이브러리(clockout-lib)의 공개 진입점 개체. 모든 Co 모듈의 공개 함수를 같은 이름의 메서드로 감싸 넘긴다.
'       Co 모듈은 Option Private Module이라 다른 통합 문서에서 직접 부를 수 없으므로 이 개체를 거친다.
' 사용법(추가기능 쪽, 늦은 바인딩):
'   Dim lib As Object
'   Set lib = Application.Run("'clockout-lib.xlam'!CoLib_Api")
'   If lib.CoRange_GetSelection(ws, rng, why) Then ...      ' ByRef 인자도 그대로 통한다
'   tok = lib.CoGuard_Begin() : ... : lib.CoGuard_End tok   ' 상태는 개체 안에 보관(토큰)
'   v = lib.CoFitMath_CalcBox(...)                          ' Variant 배열(left, top, width, height), 못 맞추면 Empty
'   s = lib.FilterPictures                                  ' 상수는 읽기 속성으로
' 주의: 이 파일은 tools/gen_colibapi.py가 만든다. 손으로 고치지 말고 lib 함수를 고친 뒤 스크립트를 다시 돌린다.
' 의존: addins/_common의 모든 Co 모듈과 ClockOutSettings.bas
'==============================================================
Option Explicit

Private mStates() As CoAppState
Private mStateCount As Long

'--- 상수 읽기 ------------------------------------------------
Public Property Get MsgTitle() As String: MsgTitle = CO_MSG_TITLE: End Property
Public Property Get FilterAll() As String: FilterAll = CO_FILTER_ALL: End Property
Public Property Get FilterPictures() As String: FilterPictures = CO_FILTER_PICTURES: End Property
Public Property Get FilterExcel() As String: FilterExcel = CO_FILTER_EXCEL: End Property
Public Property Get FilterCsv() As String: FilterCsv = CO_FILTER_CSV: End Property
Public Property Get MmToPtFactor() As Double: MmToPtFactor = CO_MM_TO_PT: End Property

'--- ClockOutSettings ----------------------------------------
Public Function ClockOut_GetSetting(ByVal addinId As String, ByVal key As String, ByVal defaultValue As Variant) As Variant
    ClockOut_GetSetting = ClockOutSettings.ClockOut_GetSetting(addinId, key, defaultValue)
End Function

'--- CoFile --------------------------------------------------
Public Function CoFile_PickFile(Optional ByVal title As String = "파일 선택", Optional ByVal fileFilter As String = CO_FILTER_ALL) As String
    CoFile_PickFile = CoFile.CoFile_PickFile(title, fileFilter)
End Function
Public Function CoFile_PickFiles(Optional ByVal title As String = "파일 선택", Optional ByVal fileFilter As String = CO_FILTER_ALL) As Collection
    Set CoFile_PickFiles = CoFile.CoFile_PickFiles(title, fileFilter)
End Function
Public Function CoFile_PickFolder(Optional ByVal title As String = "폴더 선택", Optional ByVal startFolder As String = "") As String
    CoFile_PickFolder = CoFile.CoFile_PickFolder(title, startFolder)
End Function
Public Function CoFile_PickSaveAs(Optional ByVal title As String = "저장 위치", Optional ByVal suggestedName As String = "", Optional ByVal fileFilter As String = CO_FILTER_ALL) As String
    CoFile_PickSaveAs = CoFile.CoFile_PickSaveAs(title, suggestedName, fileFilter)
End Function
Public Function CoFile_Exists(ByVal path As String) As Boolean
    CoFile_Exists = CoFile.CoFile_Exists(path)
End Function
Public Function CoFile_FolderExists(ByVal path As String) As Boolean
    CoFile_FolderExists = CoFile.CoFile_FolderExists(path)
End Function
Public Function CoFile_EnsureFolder(ByVal path As String) As Boolean
    CoFile_EnsureFolder = CoFile.CoFile_EnsureFolder(path)
End Function
Public Function CoFile_WorkbookFolder(ByVal wb As Workbook) As String
    CoFile_WorkbookFolder = CoFile.CoFile_WorkbookFolder(wb)
End Function
Public Function CoFile_ReadUtf8(ByVal path As String, ByRef text As String) As Boolean
    CoFile_ReadUtf8 = CoFile.CoFile_ReadUtf8(path, text)
End Function
Public Function CoFile_WriteUtf8(ByVal path As String, ByVal text As String, Optional ByVal withBom As Boolean = False) As Boolean
    CoFile_WriteUtf8 = CoFile.CoFile_WriteUtf8(path, text, withBom)
End Function

'--- CoFitMath -----------------------------------------------
' 결과를 Variant 배열(0:left, 1:top, 2:width, 3:height)로 돌려준다. 맞출 수 없으면 Empty.
Public Function CoFitMath_CalcBox(ByVal areaLeft As Double, ByVal areaTop As Double, _
                                  ByVal areaW As Double, ByVal areaH As Double, _
                                  ByVal marginPt As Double, _
                                  ByVal shapeW As Double, ByVal shapeH As Double, _
                                  ByVal keepRatio As Boolean, ByVal alignCenter As Boolean) As Variant
    Dim b As CoFitBox
    b = CoFitMath.CoFitMath_CalcBox(areaLeft, areaTop, areaW, areaH, marginPt, shapeW, shapeH, keepRatio, alignCenter)
    If b.Ok Then CoFitMath_CalcBox = Array(b.Left, b.Top, b.Width, b.Height)
End Function

'--- CoGuard -------------------------------------------------
' 상태(CoAppState)는 개체 경계를 넘지 못하므로 개체 안에 보관하고 토큰(1부터)을 돌려준다. CoGuard_End에 그대로 넘긴다.
Public Function CoGuard_Begin(Optional ByVal noScreen As Boolean = True, _
                              Optional ByVal noEvents As Boolean = False, _
                              Optional ByVal manualCalc As Boolean = False, _
                              Optional ByVal noAlerts As Boolean = False) As Long
    mStateCount = mStateCount + 1
    ReDim Preserve mStates(1 To mStateCount)
    mStates(mStateCount) = CoGuard.CoGuard_Begin(noScreen, noEvents, manualCalc, noAlerts)
    CoGuard_Begin = mStateCount
End Function
' CoGuard_Begin이 준 토큰의 상태로 되돌린다. 모르는 토큰이거나 이미 되돌렸으면 아무것도 안 한다.
Public Sub CoGuard_End(ByVal token As Long)
    If token < 1 Or token > mStateCount Then Exit Sub
    CoGuard.CoGuard_End mStates(token)
End Sub
Public Sub CoGuard_Status(ByVal text As String)
    CoGuard.CoGuard_Status text
End Sub
Public Function CoGuard_ErrText(Optional ByVal context As String = "작업 중") As String
    CoGuard_ErrText = CoGuard.CoGuard_ErrText(context)
End Function
Public Sub CoGuard_Notice(ByVal msg As String, Optional ByVal title As String = CO_MSG_TITLE)
    CoGuard.CoGuard_Notice msg, title
End Sub
Public Sub CoGuard_Info(ByVal msg As String, Optional ByVal title As String = CO_MSG_TITLE)
    CoGuard.CoGuard_Info msg, title
End Sub
Public Function CoGuard_Ask(ByVal msg As String, Optional ByVal title As String = CO_MSG_TITLE) As Boolean
    CoGuard_Ask = CoGuard.CoGuard_Ask(msg, title)
End Function
Public Function CoGuard_AskText(ByVal msg As String, ByRef result As String, Optional ByVal title As String = CO_MSG_TITLE, Optional ByVal defaultText As String = "") As Boolean
    CoGuard_AskText = CoGuard.CoGuard_AskText(msg, result, title, defaultText)
End Function

'--- CoNum ---------------------------------------------------
Public Function CoNum_Clamp(ByVal v As Double, ByVal lo As Double, ByVal hi As Double) As Double
    CoNum_Clamp = CoNum.CoNum_Clamp(v, lo, hi)
End Function
Public Function CoNum_IsBetween(ByVal v As Double, ByVal lo As Double, ByVal hi As Double) As Boolean
    CoNum_IsBetween = CoNum.CoNum_IsBetween(v, lo, hi)
End Function
Public Function CoNum_MmToPt(ByVal mm As Double) As Double
    CoNum_MmToPt = CoNum.CoNum_MmToPt(mm)
End Function
Public Function CoNum_CmToPt(ByVal cm As Double) As Double
    CoNum_CmToPt = CoNum.CoNum_CmToPt(cm)
End Function
Public Function CoNum_PtToMm(ByVal pt As Double) As Double
    CoNum_PtToMm = CoNum.CoNum_PtToMm(pt)
End Function
Public Function CoNum_SafeDiv(ByVal a As Double, ByVal b As Double, Optional ByVal fallback As Double = 0) As Double
    CoNum_SafeDiv = CoNum.CoNum_SafeDiv(a, b, fallback)
End Function
Public Function CoNum_Near(ByVal a As Double, ByVal b As Double, Optional ByVal eps As Double = 0.000000001) As Boolean
    CoNum_Near = CoNum.CoNum_Near(a, b, eps)
End Function
Public Function CoNum_TryParse(ByVal s As String, ByRef result As Double) As Boolean
    CoNum_TryParse = CoNum.CoNum_TryParse(s, result)
End Function
Public Function CoNum_RoundHalfUp(ByVal v As Double, Optional ByVal digits As Long = 0) As Double
    CoNum_RoundHalfUp = CoNum.CoNum_RoundHalfUp(v, digits)
End Function
Public Function CoNum_TruncTo(ByVal v As Double, ByVal unitValue As Double) As Double
    CoNum_TruncTo = CoNum.CoNum_TruncTo(v, unitValue)
End Function
Public Function CoNum_ToKoreanWon(ByVal amount As Double, Optional ByVal withPrefix As Boolean = True) As String
    CoNum_ToKoreanWon = CoNum.CoNum_ToKoreanWon(amount, withPrefix)
End Function

'--- CoRange -------------------------------------------------
Public Function CoRange_ActiveWorksheet(ByRef ws As Worksheet, ByRef reason As String) As Boolean
    CoRange_ActiveWorksheet = CoRange.CoRange_ActiveWorksheet(ws, reason)
End Function
Public Function CoRange_GetSelection(ByRef ws As Worksheet, ByRef rng As Range, ByRef reason As String, Optional ByVal allowMulti As Boolean = False, Optional ByVal allowWhole As Boolean = False) As Boolean
    CoRange_GetSelection = CoRange.CoRange_GetSelection(ws, rng, reason, allowMulti, allowWhole)
End Function
Public Function CoRange_IsWholeRowOrCol(ByVal rng As Range) As Boolean
    CoRange_IsWholeRowOrCol = CoRange.CoRange_IsWholeRowOrCol(rng)
End Function
Public Function CoRange_IsVisible(ByVal rng As Range) As Boolean
    CoRange_IsVisible = CoRange.CoRange_IsVisible(rng)
End Function
Public Function CoRange_FirstCellArea(ByVal rng As Range) As Range
    Set CoRange_FirstCellArea = CoRange.CoRange_FirstCellArea(rng)
End Function
Public Function CoRange_HasMerged(ByVal rng As Range) As Boolean
    CoRange_HasMerged = CoRange.CoRange_HasMerged(rng)
End Function
Public Function CoRange_IsProtected(ByVal ws As Worksheet, Optional ByVal forShapes As Boolean = False) As Boolean
    CoRange_IsProtected = CoRange.CoRange_IsProtected(ws, forShapes)
End Function
Public Function CoRange_CanEdit(ByVal rng As Range) As Boolean
    CoRange_CanEdit = CoRange.CoRange_CanEdit(rng)
End Function
Public Function CoRange_LastRow(ByVal ws As Worksheet, Optional ByVal col As Long = 0) As Long
    CoRange_LastRow = CoRange.CoRange_LastRow(ws, col)
End Function
Public Function CoRange_LastCol(ByVal ws As Worksheet, Optional ByVal row As Long = 0) As Long
    CoRange_LastCol = CoRange.CoRange_LastCol(ws, row)
End Function
Public Function CoRange_Values2D(ByVal rng As Range) As Variant
    CoRange_Values2D = CoRange.CoRange_Values2D(rng)
End Function

'--- CoSetting -----------------------------------------------
Public Function CoSetting_Number(ByVal addinId As String, ByVal key As String, ByVal defaultValue As Double, Optional ByVal lo As Variant, Optional ByVal hi As Variant) As Double
    CoSetting_Number = CoSetting.CoSetting_Number(addinId, key, defaultValue, lo, hi)
End Function
Public Function CoSetting_MmAsPt(ByVal addinId As String, ByVal key As String, ByVal defaultMm As Double, Optional ByVal loMm As Variant, Optional ByVal hiMm As Variant) As Double
    CoSetting_MmAsPt = CoSetting.CoSetting_MmAsPt(addinId, key, defaultMm, loMm, hiMm)
End Function
Public Function CoSetting_Bool(ByVal addinId As String, ByVal key As String, ByVal defaultValue As Boolean) As Boolean
    CoSetting_Bool = CoSetting.CoSetting_Bool(addinId, key, defaultValue)
End Function
Public Function CoSetting_Text(ByVal addinId As String, ByVal key As String, ByVal defaultValue As String, Optional ByVal trimIt As Boolean = True) As String
    CoSetting_Text = CoSetting.CoSetting_Text(addinId, key, defaultValue, trimIt)
End Function
' 선택지 설정. ParamArray는 넘겨줄 수 없어 읽은 값을 여기서 options와 견준다(동작은 CoSetting_Choice와 같다).
Public Function CoSetting_Choice(ByVal addinId As String, ByVal key As String, ByVal defaultValue As String, _
                                 ParamArray options() As Variant) As String
    Dim s As String, i As Long
    s = CoSetting.CoSetting_Choice(addinId, key, defaultValue)   ' 선택지 없이 부르면 읽은 값 그대로
    CoSetting_Choice = defaultValue
    If UBound(options) < LBound(options) Then CoSetting_Choice = s: Exit Function
    For i = LBound(options) To UBound(options)
        If s = CStr(options(i)) Then CoSetting_Choice = s: Exit Function
    Next i
End Function

'--- CoShape -------------------------------------------------
Public Function CoShape_GetSelected(ByRef sr As ShapeRange) As Boolean
    CoShape_GetSelected = CoShape.CoShape_GetSelected(sr)
End Function
Public Function CoShape_AddPicture(ByVal ws As Worksheet, ByVal filePath As String, ByRef reason As String) As Shape
    Set CoShape_AddPicture = CoShape.CoShape_AddPicture(ws, filePath, reason)
End Function
Public Function CoShape_LastAdded(ByVal ws As Worksheet, ByVal beforeCount As Long) As Shape
    Set CoShape_LastAdded = CoShape.CoShape_LastAdded(ws, beforeCount)
End Function
Public Sub CoShape_Delete(ByRef shp As Shape)
    CoShape.CoShape_Delete shp
End Sub
Public Function CoShape_IsUpright(ByVal shp As Shape) As Boolean
    CoShape_IsUpright = CoShape.CoShape_IsUpright(shp)
End Function
Public Sub CoShape_RestoreOriginalSize(ByVal shp As Shape)
    CoShape.CoShape_RestoreOriginalSize shp
End Sub
Public Function CoShape_FitToRange(ByVal shp As Shape, ByVal area As Range, ByVal marginPt As Double, Optional ByVal keepRatio As Boolean = True, Optional ByVal alignCenter As Boolean = True, Optional ByVal placement As Long = 0) As Boolean
    CoShape_FitToRange = CoShape.CoShape_FitToRange(shp, area, marginPt, keepRatio, alignCenter, placement)
End Function
Public Function CoShape_InRange(ByVal area As Range) As Collection
    Set CoShape_InRange = CoShape.CoShape_InRange(area)
End Function

'--- CoText --------------------------------------------------
Public Function CoText_IsBlank(ByVal v As Variant) As Boolean
    CoText_IsBlank = CoText.CoText_IsBlank(v)
End Function
Public Function CoText_StartsWith(ByVal s As String, ByVal prefix As String, Optional ByVal ignoreCase As Boolean = False) As Boolean
    CoText_StartsWith = CoText.CoText_StartsWith(s, prefix, ignoreCase)
End Function
Public Function CoText_EndsWith(ByVal s As String, ByVal suffix As String, Optional ByVal ignoreCase As Boolean = False) As Boolean
    CoText_EndsWith = CoText.CoText_EndsWith(s, suffix, ignoreCase)
End Function
Public Function CoText_SplitClean(ByVal s As String, Optional ByVal delim As String = ",") As Collection
    Set CoText_SplitClean = CoText.CoText_SplitClean(s, delim)
End Function
Public Function CoText_SafeFileName(ByVal fileName As String, Optional ByVal replaceWith As String = "_") As String
    CoText_SafeFileName = CoText.CoText_SafeFileName(fileName, replaceWith)
End Function
Public Function CoText_TrimSlash(ByVal path As String) As String
    CoText_TrimSlash = CoText.CoText_TrimSlash(path)
End Function
Public Function CoText_PathJoin(ByVal folder As String, ByVal child As String) As String
    CoText_PathJoin = CoText.CoText_PathJoin(folder, child)
End Function
Public Function CoText_FileName(ByVal path As String) As String
    CoText_FileName = CoText.CoText_FileName(path)
End Function
Public Function CoText_Ext(ByVal path As String) As String
    CoText_Ext = CoText.CoText_Ext(path)
End Function
Public Function CoText_BaseName(ByVal path As String) As String
    CoText_BaseName = CoText.CoText_BaseName(path)
End Function
Public Function CoText_ParentFolder(ByVal path As String) As String
    CoText_ParentFolder = CoText.CoText_ParentFolder(path)
End Function
Public Function CoText_ColLetter(ByVal n As Long) As String
    CoText_ColLetter = CoText.CoText_ColLetter(n)
End Function
Public Function CoText_ColNumber(ByVal s As String) As Long
    CoText_ColNumber = CoText.CoText_ColNumber(s)
End Function
Public Function CoText_CleanInvisible(ByVal s As String, Optional ByVal keepLineBreak As Boolean = False) As String
    CoText_CleanInvisible = CoText.CoText_CleanInvisible(s, keepLineBreak)
End Function
Public Function CoText_SafeSheetName(ByVal sheetName As String, Optional ByVal replaceWith As String = "_") As String
    CoText_SafeSheetName = CoText.CoText_SafeSheetName(sheetName, replaceWith)
End Function
Public Function CoText_NaturalCompare(ByVal a As String, ByVal b As String) As Long
    CoText_NaturalCompare = CoText.CoText_NaturalCompare(a, b)
End Function
Public Function CoText_ParseDate(ByVal s As String, ByRef result As Date) As Boolean
    CoText_ParseDate = CoText.CoText_ParseDate(s, result)
End Function
Public Function CoText_JoinList(ByVal items As Collection, ByVal delim As String, Optional ByVal prefix As String = "", Optional ByVal suffix As String = "") As String
    CoText_JoinList = CoText.CoText_JoinList(items, delim, prefix, suffix)
End Function
