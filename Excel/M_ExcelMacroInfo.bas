Attribute VB_Name = "M_ExcelMacroInfo"
Option Explicit

'Excel の Alt + F8 から表示されるマクロの情報を出力する。
'Imports VBIDE = Microsoft Visual Basic for Applications Extensibility 5.3
'Imports Scripting = Microsoft Scripting Runtime
'Imports VBScript_RegExp_55 = Microsoft VBScript Regular Expressions 5.5
Sub ListExcelMacroOptions()
    Dim targetProject As VBIDE.VBProject
    Set targetProject = Excel.Application.VBE.ActiveVBProject
    
    Dim fso As Scripting.FileSystemObject
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    Dim tmpFile As String
    tmpFile = fso.BuildPath(VBA.Environ$("TEMP"), fso.GetTempName())
    
    Dim vbc As VBIDE.VBComponent
    For Each vbc In targetProject.VBComponents
        '標準モジュールとシートモジュール以外を除く(それ以外は Alt + F8 から実行できないため)
        Select Case vbc.Type
            Case vbext_ct_StdModule, _
                 vbext_ct_Document
                'OK
                'vbext_ct_StdModule = 1
                'vbext_ct_Document  = &H64
            Case Else
                GoTo Continue
        End Select
        
        vbc.Export tmpFile
        Dim codeTxt As String
        codeTxt = fso.OpenTextFile(tmpFile, ForReading).ReadAll()
        
        PrintPublicSubProcedureInfo codeTxt
        
        fso.DeleteFile tmpFile
Continue:
    Next vbc
    
End Sub

Private Sub PrintPublicSubProcedureInfo(codeText As String)
    'Attribute VB_Name = "Module1"
    Dim moduleName As String
    moduleName = _
        NewRegExp("^Attribute VB_Name = ""(.*?)""$", False, True) _
            .Execute(codeText) _
            .Item(0) _
            .SubMatches _
            .Item(0)
    
    Dim publicSubSeacher As VBScript_RegExp_55.RegExp
    Set publicSubSeacher = NewRegExp( _
        "(?:Public (?:_\r\n)? *|^)Sub (?:_\r\n)? *(\S*?)\(", _
        GlobalMatch:=True, _
        MultiLine:=True _
    )
    
    'Sub プロシージャの辞書
    Dim subDic As Scripting.Dictionary
    Set subDic = newDic()
    
    Dim m As VBScript_RegExp_55.Match
    For Each m In publicSubSeacher.Execute(codeText)
        Set subDic.Item(m.SubMatches.Item(0)) = newDic()
    Next m
    
    
    Dim shortcutkeySearcher As VBScript_RegExp_55.RegExp
    Set shortcutkeySearcher = NewRegExp( _
        "^Attribute (.*?)\.VB_ProcData\.VB_Invoke_Func = ""(.*?)\\n14""$", _
        GlobalMatch:=True, _
        MultiLine:=True _
    )
    
    For Each m In shortcutkeySearcher.Execute(codeText)
        If subDic.Exists(m.SubMatches.Item(0)) Then
            subDic.Item(m.SubMatches.Item(0)).Item("ShortcutKey") = m.SubMatches.Item(1)
        End If
    Next m
    
    
    Dim descriptionSearcher As VBScript_RegExp_55.RegExp
    Set descriptionSearcher = NewRegExp( _
        "^Attribute (.*?)\.VB_Description = ""(.*?)""$", _
        GlobalMatch:=True, _
        MultiLine:=True _
    )
    
    For Each m In descriptionSearcher.Execute(codeText)
        If subDic.Exists(m.SubMatches.Item(0)) Then
            subDic.Item(m.SubMatches.Item(0)).Item("Description") = m.SubMatches.Item(1)
        End If
    Next m
    
    
    'Sub プロシージャ一覧
    Dim k As Variant
    For Each k In subDic.Keys()
        Dim macroName As String
        macroName = moduleName & "." & k
        
        Dim dic As Scripting.Dictionary
        Set dic = subDic.Item(k)
        
        Dim macroShortcutKey As String
        macroShortcutKey = dic.Item("ShortcutKey")
        
        Dim macroDescription As String
        macroDescription = dic.Item("Description")
        
        Debug.Print macroName, macroShortcutKey, macroDescription
    Next k
End Sub

Private Function newDic() As Scripting.Dictionary
    Set newDic = VBA.CreateObject("Scripting.Dictionary")
End Function

Private Function NewRegExp( _
        Pattern As String, _
        Optional GlobalMatch As Boolean, _
        Optional MultiLine As Boolean, _
        Optional IgnoreCase As Boolean _
    ) As VBScript_RegExp_55.RegExp
    
    Dim regEx As VBScript_RegExp_55.RegExp
    Set regEx = VBA.CreateObject("VBScript.RegExp")
    With regEx
        .Pattern = Pattern
        .Global = GlobalMatch
        .MultiLine = MultiLine
        .IgnoreCase = IgnoreCase
    End With
    
    Set NewRegExp = regEx
    
End Function
