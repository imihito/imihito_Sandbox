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
    
    'JScript を 64bit の VBA からも使用できるようにしたクラス。
    'MSHTML.HTMLDocument 上で JScript を実行することで ScriptContorl の 32bit 制限を回避する。
Public aaaaa
Option Explicit

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
    '`inFunctionBody`内で使用する引数。
    '複数指定時はカンマ区切りで指定する。
    '参考:'[Function - JavaScript | MDN](https://developer.mozilla.org/ja/docs/Web/JavaScript/Reference/Global_Objects/Function)
'inFunctionBody
    '関数本文。
'inAutoReturn
    '省略可能。省略時True。
    'Trueのとき`inFunctionBody`の先頭に`return `を追加する。

'Return
    'インスタンスされたJScriptのfunctionオブジェクト。

Public Function Func( _
                 inArguments As String, _
                 inFunctionBody As String, _
        Optional inAutoReturn As Boolean = True _
    ) As Object
    
    Const EXEC_SCRIPT = _
            "this.createFunc=" & _
                "function(args,funcBody){" & _
                    "return new Function(args,funcBody);}"
    
    '各種初期化
    '関数オブジェクトのキャッシュ
    Static funcCache As Object 'As Scripting.Dictionary
    If funcCache Is Nothing Then _
        Set funcCache = VBA.Interaction.CreateObject("Scripting.Dictionary")
    
    'JScript実行環境。参照を保持しないとインスタンスしたfunctionオブジェクトも消える
    Static htmlDoc    As Object 'As MSHTML.HTMLDocument
    Static createFunc As Object 'JScript function
    
    If htmlDoc Is Nothing Then
        Call funcCache.RemoveAll
        Set htmlDoc = VBA.Interaction.CreateObject("htmlfile")
        
        'JScriptのグローバル変数に関数を定義
        Call htmlDoc.parentWindow.execScript(EXEC_SCRIPT)
        
        '作成した関数を静的変数に保管（書き換え防止）
        Set createFunc = htmlDoc.parentWindow.createFunc
    End If


    'キャッシュ用に整形
    Dim trimedArgs As String, trimedBody As String
    trimedArgs = VBA.Strings.Trim$(inArguments)
    If inAutoReturn Then
        trimedBody = "return " & VBA.Strings.Trim$(inFunctionBody)
    Else
        trimedBody = VBA.Strings.Trim$(inFunctionBody)
    End If

    Dim cacheKey As String
    cacheKey = trimedArgs & "|" & trimedBody


    'キャッシュに無ければ新規インスタンス
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



'汎用代入関数。
Private Property Let SetVar( _
        ByRef outVariable As Variant, _
              inExpression As Variant _
    )
    '判定の順番は変更不可。
    If VBA.Information.IsObject(inExpression) Then
        Set outVariable = inExpression
    ElseIf VBA.Information.VarType(inExpression) = VBA.VbVarType.vbDataObject Then
        Set outVariable = inExpression
    Else
        Let outVariable = inExpression
    End If
End Property

