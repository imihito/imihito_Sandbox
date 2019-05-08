Attribute VB_Name = "M_ChangeOneNoteFont"
Option Explicit

'以下の参照設定は必須。
'OneNote = Microsoft OneNote 15.0 Object Library

Public Sub ChangeOneSectionFont()
    'OneNote 取得。
    Dim appOne As OneNote.Application
    Set appOne = VBA.CreateObject("OneNote.Application")
    
    
    '最前面のセクションの ID を取得。
        'appOne の型を明示しないと以下のエラーが発生する。
            'オートメーション エラーです。
            'ライブラリは登録されていません｡
    Dim sectId As String
    sectId = appOne.Windows.CurrentWindow.CurrentSectionId
    
    'セクションのページを示した XML 文字列を取得。
    Dim hierarchyXml As String
    appOne.GetHierarchy sectId, OneNote.HierarchyScope.hsPages, hierarchyXml
    
    'XML DOM としてロード。
    Dim hierarchyXmlDoc As Object 'As MSXML2.DOMDocument
    Set hierarchyXmlDoc = newXmlDoc(hierarchyXml)
    
    'ページを取得し、それぞれのフォントを変更。
    Dim node As Object 'As MSXML2.IXMLDOMNode
    For Each node In hierarchyXmlDoc.getElementsByTagName("one:Page")
        ChangeOnePageFont appOne, node.Attributes.getNamedItem("ID").NodeValue
    Next node
    
End Sub


Public Sub ChangeOnePageFont( _
        inAppOne As OneNote.Application, _
        inPageId As String _
    )
    
    'ページの内容を取得。
    Dim contentsBuf As String
    inAppOne.GetPageContent inPageId, contentsBuf
    
    'XML DOM としてロード。
    Dim pageXml As Object 'As MSXML2.DOMDocument
    Set pageXml = newXmlDoc(contentsBuf)
    
    'スタイルのフォントを変更。
    Const BaseFont = "游ゴシック"
    Const NewFont = "Meiryo UI"
    
    'BaseFont が指定されているスタイルについてフォントを変更する。
    Dim node As Object 'As MSXML2.IXMLDOMNode
    For Each node In pageXml.SelectNodes("//one:QuickStyleDef[@font='" & BaseFont & "']")
        node.Attributes.getNamedItem("font").NodeValue = NewFont
    Next node
    
    
    '変更内容を書き戻し。
    inAppOne.UpdatePageContent pageXml.XML
    
End Sub

'MSXML2 = Microsoft XML, v3.0
Private Function newXmlDoc(inXmlString As String) As Object 'As MSXML2.DOMDocument
    Dim xmlDoc As Object 'As MSXML2.DOMDocument
    Set xmlDoc = VBA.CreateObject("MSXML2.DOMDocument")
    xmlDoc.LoadXML inXmlString
    Set newXmlDoc = xmlDoc
End Function
