VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JScriptEnvironment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'Instance JScript Function object

'e.g.
    'Dim adder As Object: Set adder = JSFunc("a,b", "a+b") 'autoReturn = True
    'Debug.Print adder(2, 6) '->8

    'Dim inRange As Object
    'Set inRange = JSFunc("range,min,max", "v=range.Value;return min<=v&&v<=max;", False) 'autoReturn = False
    'Excel.ActiveCell.Value() = 150
    'Debug.Print inRange(Excel.ActiveCell, 100, 200) '->True

'Arguments
'args
    '`funcBody`内で使用する引数。
    '複数指定時はカンマ区切りで指定する。
    '参考:'[Function - JavaScript | MDN](https://developer.mozilla.org/ja/docs/Web/JavaScript/Reference/Global_Objects/Function)
'funcBody
    '関数本文。
'autoReturn
    '省略可能。省略時True。
    'Trueのとき`funcBody`の先頭に`return `を追加する。

'Return
    'インスタンスされたJScriptのfunctionオブジェクト。

Public Function NewFunc( _
        args As String, _
        funcBody As String, _
        Optional autoReturn As Boolean = True) As Object

    Const EXEC_SCRIPT = _
            "this.createFunc=" & _
                "function(args,funcBody){" & _
                    "return new Function(args,funcBody);}"

    '各種初期化
    '関数オブジェクトのキャッシュ
    Static funcCache As Object 'As Scripting.Dictionary
    If funcCache Is Nothing Then Set funcCache = VBA.CreateObject("Scripting.Dictionary")

    'JScript実行環境。参照を保持しないとインスタンスしたfunctionオブジェクトも消える
    Static htmlDoc    As Object 'As MSHTML.HTMLDocument
    Static createFunc As Object 'JScript function
    
    If htmlDoc Is Nothing Then
        Call funcCache.RemoveAll
        Set htmlDoc = VBA.CreateObject("htmlfile")
        
        'JScriptのグローバル変数に関数を定義
        Call htmlDoc.parentWindow.execScript(EXEC_SCRIPT)
        
        '作成した関数を静的変数に保管（書き換え防止）
        Set createFunc = htmlDoc.parentWindow.createFunc
    End If


    'キャッシュ用に整形
    Dim trimedArgs As String, trimedBody As String
    trimedArgs = VBA.Trim$(args)
    If autoReturn Then
        trimedBody = "return " & VBA.Trim$(funcBody)
    Else
        trimedBody = VBA.Trim$(funcBody)
    End If

    Dim cacheKey As String
    cacheKey = trimedArgs & "|" & trimedBody


    'キャッシュに無ければ新規インスタンス
    If Not funcCache.Exists(cacheKey) Then
        Call funcCache.Add(cacheKey, createFunc(trimedArgs, trimedBody))
    End If

    Set NewFunc = funcCache.Item(cacheKey)

End Function


Public Function Parse(jsExpression As String) As Variant
    Dim evalFunc As Object
    Set evalFunc = Me.NewFunc("s", "eval('('+s+')')")
    SetVar(Parse) = evalFunc(jsExpression)
End Function


Public Function IsJsObject(ByVal iObject As Object) As Boolean
    Let IsJsObject = Me.NewFunc("o", "o instanceof Object")(iObject)
End Function



Private Property Let SetVar( _
        ByRef outVariable As Variant, _
        inExpression As Variant _
    )
    Select Case True
        Case VBA.IsObject(inExpression), _
             VBA.VarType(inExpression) = vbDataObject
            
            Set outVariable = inExpression
            
        Case Else
            Let outVariable = inExpression
            
    End Select
End Property

