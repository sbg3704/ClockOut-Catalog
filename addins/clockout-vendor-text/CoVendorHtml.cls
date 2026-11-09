VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoVendorHtml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================
' 기능: stdHTML을 만드는 팩토리. 다른 통합 문서에서 늦은 바인딩으로 stdHTML.Create를 부르며 선택 인자를 빼면
'       빠진 attributes가 기본값(Empty)으로 오지 않아 오류가 난다(2026-11-09 시험). 여기서 빠진 인자를 채워 넘긴다.
' 사용법:
'   Set h = v.CoVendor_Html().Create("p")                          ' 속성 없음
'   Set h = v.CoVendor_Html().Create("a", Array("href", "x.html"))  ' 속성은 이름·값 짝 배열
'   h.CreateLiteral "글": s = h.ToString()
'   Set h = v.CoVendor_Html().CreateFromHTML("<p>글</p>")
' 의존: stdHTML.cls
'==============================================================
Option Explicit

' 태그 하나를 만든다. attributes는 Array("이름", "값", ...) 짝 배열(없으면 생략).
Public Function Create(Optional ByVal tagName As String = "html", Optional ByVal attributes As Variant, _
                       Optional ByVal isSelfClosing As Boolean = False) As Object
    If IsMissing(attributes) Then attributes = Empty
    Set Create = stdHTML.Create(tagName, attributes, isSelfClosing)
End Function

' HTML 글을 읽어 stdHTML 개체로.
Public Function CreateFromHTML(ByVal html As String) As Object
    Set CreateFromHTML = stdHTML.CreateFromHTML(html)
End Function
