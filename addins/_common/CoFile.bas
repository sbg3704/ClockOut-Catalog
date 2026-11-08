Attribute VB_Name = "CoFile"
'==============================================================
' 기능: 칼퇴 공용 라이브러리 - 파일·폴더 고르기 창, 있는지 확인, 폴더 만들기, UTF-8 읽기·쓰기.
'       고르기 창은 취소하면 빈 글자("")를 돌려준다(GetOpenFilename의 False 비교 실수 방지).
' 사용법:
'   p = CoFile_PickFile("사진 선택", CO_FILTER_PICTURES)   : If Len(p) = 0 Then Exit Sub
'   Set files = CoFile_PickFiles("여러 파일 선택")         ' Collection, 취소면 0개
'   d = CoFile_PickFolder("저장할 폴더")
'   p = CoFile_PickSaveAs("저장 위치", "결과.csv", CO_FILTER_CSV)
'   If CoFile_ReadUtf8(p, txt) Then ...                  ' 실패하면 False, 오류를 던지지 않음
' 의존: CoText.bas(경로 나누기). Windows(ADODB.Stream 후기 바인딩, 참조 추가 불필요).
' 주의: 추가기능 파일 자체나 칼퇴 설정 파일에는 쓰지 않는다(docs/03 4번). 사용자 결과물용이다.
'==============================================================
Option Explicit
Option Private Module

Public Const CO_FILTER_ALL As String = "모든 파일 (*.*),*.*"
Public Const CO_FILTER_PICTURES As String = _
    "사진 파일 (*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tif;*.tiff),*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tif;*.tiff"
Public Const CO_FILTER_EXCEL As String = "Excel 파일 (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls"
Public Const CO_FILTER_CSV As String = "CSV 파일 (*.csv),*.csv"

Private Const DLG_FOLDER_PICKER As Long = 4           ' msoFileDialogFolderPicker

' 파일 하나를 고른다. 취소하면 "".
Public Function CoFile_PickFile(Optional ByVal title As String = "파일 선택", _
                                Optional ByVal fileFilter As String = CO_FILTER_ALL) As String
    Dim v As Variant
    v = Application.GetOpenFilename(FileFilter:=fileFilter, Title:=title, MultiSelect:=False)
    If VarType(v) = vbBoolean Then Exit Function       ' 취소(False)
    CoFile_PickFile = CStr(v)
End Function

' 여러 파일을 고른다. 고른 경로(String)들의 Collection, 취소하면 빈 Collection.
Public Function CoFile_PickFiles(Optional ByVal title As String = "파일 선택", _
                                 Optional ByVal fileFilter As String = CO_FILTER_ALL) As Collection
    Dim v As Variant, i As Long, col As New Collection
    v = Application.GetOpenFilename(FileFilter:=fileFilter, Title:=title, MultiSelect:=True)
    If IsArray(v) Then
        For i = LBound(v) To UBound(v)
            col.Add CStr(v(i))
        Next i
    End If
    Set CoFile_PickFiles = col
End Function

' 폴더를 고른다. 취소하면 "". 돌려주는 경로 끝에는 \가 없다.
Public Function CoFile_PickFolder(Optional ByVal title As String = "폴더 선택", _
                                  Optional ByVal startFolder As String = "") As String
    Dim dlg As Object
    On Error GoTo Fail
    Set dlg = Application.FileDialog(DLG_FOLDER_PICKER)
    dlg.Title = title
    dlg.AllowMultiSelect = False
    If Len(startFolder) > 0 Then dlg.InitialFileName = CoText_PathJoin(startFolder, "")
    If dlg.Show = -1 Then CoFile_PickFolder = CoText_TrimSlash(CStr(dlg.SelectedItems(1)))
    Exit Function
Fail:
    CoFile_PickFolder = ""
End Function

' 저장할 위치를 고른다(파일은 만들지 않는다). 취소하면 "".
Public Function CoFile_PickSaveAs(Optional ByVal title As String = "저장 위치", _
                                  Optional ByVal suggestedName As String = "", _
                                  Optional ByVal fileFilter As String = CO_FILTER_ALL) As String
    Dim v As Variant
    v = Application.GetSaveAsFilename(InitialFileName:=suggestedName, FileFilter:=fileFilter, Title:=title)
    If VarType(v) = vbBoolean Then Exit Function
    CoFile_PickSaveAs = CStr(v)
End Function

' 파일이 있으면 True(폴더는 False). Dir를 쓰지 않아 호출한 쪽의 Dir 반복을 깨지 않는다.
Public Function CoFile_Exists(ByVal path As String) As Boolean
    Dim a As Long
    If Len(path) = 0 Then Exit Function
    On Error GoTo Fail
    a = GetAttr(path)
    CoFile_Exists = ((a And vbDirectory) = 0)
Fail:
End Function

Public Function CoFile_FolderExists(ByVal path As String) As Boolean
    Dim a As Long
    If Len(path) = 0 Then Exit Function
    On Error GoTo Fail
    a = GetAttr(CoText_TrimSlash(path))
    CoFile_FolderExists = ((a And vbDirectory) <> 0)
Fail:
End Function

' 폴더를 (상위 폴더까지) 만든다. 이미 있거나 만들었으면 True.
Public Function CoFile_EnsureFolder(ByVal path As String) As Boolean
    Dim p As String, parent As String
    p = CoText_TrimSlash(path)
    If Len(p) = 0 Then Exit Function
    If CoFile_FolderExists(p) Then CoFile_EnsureFolder = True: Exit Function
    parent = CoText_ParentFolder(p)
    If Len(parent) > 0 And parent <> p Then
        If Not CoFile_EnsureFolder(parent) Then Exit Function
    End If
    On Error Resume Next
    MkDir p
    On Error GoTo 0
    CoFile_EnsureFolder = CoFile_FolderExists(p)
End Function

' 통합 문서가 저장된 폴더. 아직 저장 안 한 새 문서면 "".
Public Function CoFile_WorkbookFolder(ByVal wb As Workbook) As String
    On Error Resume Next
    CoFile_WorkbookFolder = wb.Path
End Function

' UTF-8 글 파일을 읽는다(BOM 있어도 됨). 실패하면 False.
Public Function CoFile_ReadUtf8(ByVal path As String, ByRef text As String) As Boolean
    Dim st As Object
    text = ""
    On Error GoTo Fail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                                       ' 글자
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    text = st.ReadText
    st.Close
    If Len(text) > 0 Then
        If AscW(Left$(text, 1)) = &HFEFF Then text = Mid$(text, 2)
    End If
    CoFile_ReadUtf8 = True
    Exit Function
Fail:
    text = ""
End Function

' 글을 UTF-8 파일로 쓴다(있으면 덮어씀). 기본은 BOM 없이. 실패하면 False.
' Excel에서 바로 열 CSV는 withBom=True로 써야 한글이 깨지지 않는다.
Public Function CoFile_WriteUtf8(ByVal path As String, ByVal text As String, _
                                 Optional ByVal withBom As Boolean = False) As Boolean
    Dim st As Object, bin As Object
    On Error GoTo Fail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText text                                 ' ADODB는 맨 앞에 BOM 3바이트를 붙인다
    If withBom Then
        st.SaveToFile path, 2                         ' 2 = 덮어쓰기
    Else
        st.Position = 0
        st.Type = 1                                   ' 바이트로 바꿔 BOM을 건너뛴다
        st.Position = 3
        Set bin = CreateObject("ADODB.Stream")
        bin.Type = 1
        bin.Open
        st.CopyTo bin
        bin.SaveToFile path, 2
        bin.Close
    End If
    st.Close
    CoFile_WriteUtf8 = True
    Exit Function
Fail:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    If Not bin Is Nothing Then bin.Close
End Function
