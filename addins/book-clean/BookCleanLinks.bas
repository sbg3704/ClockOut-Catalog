Attribute VB_Name = "BookCleanLinks"
'==============================================================
' 기능: 파일을 열 때 "링크를 업데이트하시겠습니까?"가 뜨게 하는 외부 링크, #REF!로 깨진 이름(이름 관리자)을 찾아 정리한다.
' 사용법: 시트 탭 우클릭 > 칼퇴 문서 정리 > 외부 링크·깨진 이름 정리. 찾은 목록을 보여 준 뒤 항목별로 확인받는다.
' 설정: linksBreakAction(외부 링크 처리: "값으로 바꾸고 끊기" 또는 "건드리지 않기", 기본 앞의 것).
' 규칙: 링크를 끊으면 그 링크를 쓰던 수식은 현재 값으로 굳는다(되돌릴 수 없음). 깨진 이름은 이름만 지운다.
'       수식 안에서 쓰는 깨진 이름이 있으면 그 칸은 #NAME? 오류 그대로이므로 지운 뒤 확인이 필요하다고 안내한다.
' 의존: BookClean.bas, 공용 라이브러리(CoSetting_Choice, CoGuard_*)
'==============================================================
Option Explicit
Option Private Module

Public Sub BookClean_Links()
    Dim api As Object, wb As Workbook, src As Variant, i As Long, nm As Object, broken As Collection
    Dim linkTxt As String, nameTxt As String, doLinks As Boolean, tok As Long, msg As String, nl As Long, nb As Long
    Set api = BC_Api()
    If api Is Nothing Then Exit Sub
    Set wb = ActiveWorkbook
    If wb Is Nothing Then api.CoGuard_Notice "열린 통합 문서가 없습니다.": Exit Sub
    doLinks = (api.CoSetting_Choice(BC_ADDIN_ID, "linksBreakAction", "값으로 바꾸고 끊기", "값으로 바꾸고 끊기", "건드리지 않기") = "값으로 바꾸고 끊기")
    src = wb.LinkSources(xlExcelLinks)
    If IsArray(src) Then
        nl = UBound(src) - LBound(src) + 1
        For i = LBound(src) To UBound(src)
            If i - LBound(src) < 8 Then linkTxt = linkTxt & vbCrLf & "  - " & src(i)
        Next i
        If nl > 8 Then linkTxt = linkTxt & vbCrLf & "  ... 외 " & (nl - 8) & "개"
    End If
    Set broken = New Collection
    For Each nm In wb.Names
        If InStr(1, nm.RefersTo, "#REF!", vbTextCompare) > 0 Then broken.Add nm
    Next nm
    nb = broken.Count
    For i = 1 To broken.Count
        If i <= 8 Then nameTxt = nameTxt & vbCrLf & "  - " & broken(i).Name
    Next i
    If nb > 8 Then nameTxt = nameTxt & vbCrLf & "  ... 외 " & (nb - 8) & "개"
    If nl = 0 And nb = 0 Then api.CoGuard_Notice "외부 링크도, 깨진 이름도 없습니다.": Exit Sub
    msg = ""
    If nl > 0 Then msg = msg & "외부 링크 " & nl & "개:" & linkTxt & vbCrLf & vbCrLf
    If nb > 0 Then msg = msg & "깨진 이름(#REF!) " & nb & "개:" & nameTxt & vbCrLf & vbCrLf
    If nl > 0 And Not doLinks Then msg = msg & "(외부 링크는 설정에 따라 건드리지 않습니다.)" & vbCrLf
    If Not (nb > 0 Or (nl > 0 And doLinks)) Then api.CoGuard_Notice Trim$(msg): Exit Sub
    If Not BC_Confirm(api, msg & IIf(nl > 0 And doLinks, "링크는 끊고 값으로 굳힙니다. ", "") & IIf(nb > 0, "깨진 이름은 삭제합니다.", "")) Then Exit Sub
    tok = api.CoGuard_Begin(True, True, False, True)
    On Error GoTo EH
    If nl > 0 And doLinks Then
        For i = LBound(src) To UBound(src)
            wb.BreakLink src(i), xlLinkTypeExcelLinks
        Next i
    End If
    For i = 1 To nb
        broken(i).Delete
    Next i
    api.CoGuard_End tok
    api.CoGuard_Notice "정리했습니다: 링크 " & IIf(doLinks, nl, 0) & "개 끊음, 깨진 이름 " & nb & "개 삭제." & vbCrLf & _
                       "수식에 #NAME? 오류가 보이면 지운 이름을 쓰던 칸입니다."
    Exit Sub
EH:
    msg = "정리하는 중 문제가 생겼습니다: " & Err.Description
    Err.Clear
    api.CoGuard_End tok
    api.CoGuard_Notice msg
End Sub
