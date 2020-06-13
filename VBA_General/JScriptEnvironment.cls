VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JScriptEnvironment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'JScript Environment for VBA (Support 64bit).
    'update 2020.03.13
    
    'JScript �� 64bit �� VBA ������g�p�ł���悤�ɂ����N���X�B
    'MSHTML.HTMLDocument ��� JScript �����s���邱�Ƃ� ScriptContorl �� 32bit �������������B

Option Explicit

Private rootHtmlDoc As MSHTML.HTMLDocument
Private ie10SupportDoc_ As MSHTML.HTMLDocument


'Instance JScript Function object.

'e.g.
    'Dim adder As Object: Set adder = JScriptEnvironment.Func("a,b", "a+b") 'inAutoReturn = True
    'Debug.Print adder(2, 6) '->8

    'Dim inRange As Object
    'Set inRange = JScriptEnvironment.Func("range,min,max", "v=range.Value;return min<=v&&v<=max;", False) 'inAutoReturn = False
    'Excel.ActiveCell.Value() = 150
    'Debug.Print inRange(Excel.ActiveCell, 100, 200) '->True

'Arguments
'inArguments
    '`inFunctionBody`���Ŏg�p��������B
    '�����w�莞�̓J���}��؂�Ŏw�肷��B
    '�Q�l:'[Function - JavaScript | MDN](https://developer.mozilla.org/ja/docs/Web/JavaScript/Reference/Global_Objects/Function)
'inFunctionBody
    '�֐��{���B
'inAutoReturn
    '�ȗ��\�B�ȗ���True�B
    'True�̂Ƃ�`inFunctionBody`�̐擪��`return `��ǉ�����B

'Return
    '�C���X�^���X���ꂽJScript��function�I�u�W�F�N�g�B

Public Function Func( _
                 inArguments As String, _
                 inFunctionBody As String, _
        Optional inAutoReturn As Boolean = True _
    ) As Object
    
    Const EXEC_SCRIPT = _
            "this.document.createFunc=" & _
                "function(args,funcBody){" & _
                    "return new Function(args,funcBody);}"
    
    '�e�평����
    '�֐��I�u�W�F�N�g�̃L���b�V��
    Static funcCache As Object 'As Scripting.Dictionary
    If funcCache Is Nothing Then
        Set funcCache = VBA.Interaction.CreateObject("Scripting.Dictionary")
    End If
    
    'JScript���s���B�Q�Ƃ�ێ����Ȃ��ƃC���X�^���X����function�I�u�W�F�N�g��������
    Static htmlDoc    As Object 'As MSHTML.HTMLDocument
    Static createFunc As Object 'JScript function
    
    If htmlDoc Is Nothing Then
        Call funcCache.RemoveAll
        Set htmlDoc = VBA.Interaction.CreateObject("htmlfile")
        Set ie10SupportDoc_ = createIE10SupportHtmlDoc(htmlDoc)
        'JScript�̃O���[�o���ϐ��Ɋ֐����`
        Call ie10SupportDoc_.parentWindow.execScript(EXEC_SCRIPT)
        
        '�쐬�����֐���ÓI�ϐ��ɕۊǁi���������h�~�j
        Set createFunc = ie10SupportDoc_.createFunc
    End If


    '�L���b�V���p�ɐ��`
    Dim trimedArgs As String, trimedBody As String
    trimedArgs = VBA.Strings.Trim$(inArguments)
    If inAutoReturn Then
        trimedBody = "return " & VBA.Strings.Trim$(inFunctionBody)
    Else
        trimedBody = VBA.Strings.Trim$(inFunctionBody)
    End If

    Dim cacheKey As String
    cacheKey = trimedArgs & "|" & trimedBody


    '�L���b�V���ɖ�����ΐV�K�C���X�^���X
    If Not funcCache.Exists(cacheKey) Then
        Call funcCache.Add(cacheKey, createFunc(trimedArgs, trimedBody))
    End If

    Set Func = funcCache.Item(cacheKey)

End Function


Public Function Parse(jsExpression As String) As Variant
    Dim evalFunc As Object
    Set evalFunc = Me.Func("s", "eval('('+s+')')")
    SetVar(Parse) = evalFunc(jsExpression)
End Function


Public Function IsJsObject(ByVal iObject As Object) As Boolean
    Let IsJsObject = Me.Func("o", "o instanceof Object")(iObject)
End Function


Private Function createIE10SupportHtmlDoc(inParentHtmlDoc As MSHTML.HTMLDocument) As MSHTML.HTMLDocument
    Const IE10HtmlSource = _
        "<!DOCTYPE html>" & vbCrLf & _
        "<!-- saved from url=(0017)http://localhost/ -->" & vbCrLf & _
        "<html>" & _
            "<head>" & _
                "<meta http-equiv='x-ua-compatible' content='IE=10' />" & _
                "<script type='text/vbscript'>" & _
                    "Option Explicit" & vbCrLf & _
                    "Public Function vbTypeName(a)" & vbCrLf & _
                        "vbTypeName = TypeName(a)" & vbCrLf & _
                    "End Function" & _
                "</script>" & _
            "</head>" & _
            "<body>" & _
            "</body>" & _
        "</html>"
    
    Dim fso As Scripting.FileSystemObject
    Set fso = VBA.Interaction.CreateObject("Scripting.FileSystemObject")
    
    Dim tmpHtmlPath As String
    tmpHtmlPath = generateTempFilePath("html")
    
    With fso.CreateTextFile(tmpHtmlPath, Overwrite:=False)
        .WriteLine IE10HtmlSource
        .Close
    End With
    
    Set createIE10SupportHtmlDoc = LoadHtmlDocumentFormUrl(tmpHtmlPath, inParentHtmlDoc)
    
    fso.DeleteFile tmpHtmlPath
    
End Function


Private Function LoadHtmlDocumentFormUrl(inUrl As String, inParentHtmlDoc As MSHTML.HTMLDocument) As MSHTML.HTMLDocument
    Dim doc As MSHTML.HTMLDocument
    Set doc = inParentHtmlDoc.createDocumentFromUrl(inUrl, VBA.Constants.vbNullString) 'https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752523(v%3Dvs.85)
    
    Const ExpectReadyState = "complete"
    Dim i As Long
    For i = 1 To 10 ^ 6
        If doc.readyState = ExpectReadyState Then
            Set LoadHtmlDocumentFormUrl = doc
            Exit Function
        End If
        Call VBA.Interaction.DoEvents
    Next i
    
    Err.Raise 5, , "Can't load complete document from " & inUrl
End Function


Private Function generateTempFilePath(inExt As String) As String
    Dim fso As Scripting.FileSystemObject
    Set fso = VBA.Interaction.CreateObject("Scripting.FileSystemObject")
    
    Const fsTemporaryFolder = 2
    Dim tmpFolderPath As String
    tmpFolderPath = fso.GetSpecialFolder(fsTemporaryFolder).Path
    
    Dim i As Long
    For i = 1 To 100000
        Dim tmpFilePath As String
        tmpFilePath = fso.BuildPath(tmpFolderPath, fso.GetTempName() & "." & inExt)
        Select Case True
            Case fso.FileExists(tmpFilePath), fso.FolderExists(tmpFilePath)
                'Next
            Case Else 'Not exists.
                Let generateTempFilePath = tmpFilePath
                Exit Function
        End Select
    Next i
    Err.Raise 735
End Function

'�ėp����֐��B
Private Property Let SetVar( _
        ByRef outVariable As Variant, _
              inExpression As Variant _
    )
    '����̏��Ԃ͕ύX�s�B
    If VBA.Information.IsObject(inExpression) Then
        Set outVariable = inExpression
    ElseIf VBA.Information.VarType(inExpression) = VBA.VbVarType.vbDataObject Then
        Set outVariable = inExpression
    Else
        Let outVariable = inExpression
    End If
End Property
