Attribute VB_Name = "DisableExcelGlobal"
'Excelのグローバルなメンバーを無効化するモジュール
'https://qiita.com/nukie_53/items/4e9226ac8f747e772afd

Option Explicit
Option Private Module

'親を明示しないと危ないもの、使用頻度が低そうなメンバー(主観)の定義を上書き(無効化)。
'上位階層のプロパティや関数的なメソッドのみが残る。
'Excel内で使う用。
#Const DisableForInternal = True

'ほとんどのメンバーの定義を上書き(無効化)。
'Excel以外のVBAホストからExcelを操作するとき用。
#Const DisableForExternal = True

'Public Sub Application(): End Sub

#If DisableForInternal Or DisableForExternal Then

'Method override
Public Sub Calculate(a): End Sub
Public Sub DDEExecute(): End Sub
Public Sub DDEInitiate(): End Sub
Public Sub DDEPoke(): End Sub
Public Sub DDERequest(): End Sub
Public Sub DDETerminate(): End Sub

'Property override
Public Sub ActiveCell(): End Sub
Public Sub ActiveChart(): End Sub
Public Sub ActivePrinter(): End Sub
Public Sub ActiveSheet(): End Sub
Public Sub ActiveWindow(): End Sub
Public Sub ActiveWorkbook(): End Sub
Public Sub AddIns(): End Sub
Public Sub Cells(): End Sub
Public Sub Charts(): End Sub
Public Sub Columns(): End Sub
Public Sub CommandBars(): End Sub
Public Sub Creator(): End Sub
Public Sub DDEAppReturnCode(): End Sub
Public Sub Excel4IntlMacroSheets(): End Sub
Public Sub Excel4MacroSheets(): End Sub
Public Sub Names(): End Sub
Public Sub Parent(): End Sub
Public Sub Range(): End Sub
Public Sub Rows(): End Sub
Public Sub Selection(): End Sub
Public Sub Sheets(): End Sub
Public Sub Worksheets(): End Sub

'Hidden Property
Public Sub ActiveDialog(): End Sub
Public Sub ActiveMenuBar(): End Sub
Public Sub Assistant(): End Sub
Public Sub DialogSheets(): End Sub
Public Sub MenuBars(): End Sub
Public Sub Modules(): End Sub
Public Sub ShortcutMenus(): End Sub
Public Sub Toolbars(): End Sub
#End If


#If DisableForExternal Then
'最上位クラスのプロパティや関数的に使えるメソッド群

'Method override
Public Sub Evaluate(): End Sub
Public Sub ExecuteExcel4Macro(): End Sub
Public Sub Intersect(): End Sub
Public Sub Run(a, b, c, d, e, f, g, h, i, j, k, l, m, n, o, p, q, r, s, t, u, v, w, x, y, z): End Sub
'Public Sub SendKeys(): End Sub 'VBA has SendKeys method.
Public Sub Union(): End Sub

'Property override
Public Sub ThisWorkbook(): End Sub
Public Sub Windows(): End Sub
Public Sub Workbooks(): End Sub
Public Sub WorksheetFunction(): End Sub
#End If
