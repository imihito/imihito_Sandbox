Attribute VB_Name = "WinAPIforDialogs"
Option Explicit

'TaskDialog の引数用の列挙型。
    '参考:https://memos-by-oxalis.hatenablog.com/entry/2019/09/22/223121
Public Enum TaskDialogButtons
    tdOK = 2 ^ 0
    tdYes = 2 ^ 1
    tdNo = 2 ^ 2
    tdCancel = 2 ^ 3 '設定しないと×ボタンも表示されない。
    tdRetry = 2 ^ 4
    tdClose = 2 ^ 5
    
    tdOKCancel = tdOK Or tdCancel
    tdRetryCancel = tdRetry Or tdCancel
    tdYesNo = tdYes Or tdNo
    tdYesNoCancel = tdYes Or tdNo Or tdCancel
End Enum

Public Enum TaskDialogIcons
    tdNone = 0&
    tdWarning = &HFFFF&     '黄色の△!  vbExclamation とかと同じ。
    tdError = &HFFFE&       '赤色の○× vbCritical とかと同じ。
    tdInformation = &HFFFD& '青色の○i  vbInformation とかと同じ。
    tdShield = &HFFFC&
End Enum

Private Enum HRESUT
    S_OK = 0&
End Enum

Private Declare PtrSafe Function _
    TaskDialog Lib "Comctl32.dll" ( _
        ByVal hwndOwner As LongPtr, _
        ByVal hInstance As LongPtr, _
        ByVal pszWindowTitle As LongPtr, _
        ByVal pszMainInstruction As LongPtr, _
        ByVal pszContent As LongPtr, _
        ByVal dwCommonButtons As TaskDialogButtons, _
        ByVal pszIcon As LongPtr, _
        ByRef pnButton As VBA.VbMsgBoxResult _
    ) As HRESUT

'今のプロセスで最前面にあるウィンドウのハンドルを取得する。
Private Declare PtrSafe Function _
    GetActiveWindow Lib "User32.dll" ( _
    ) As LongPtr

Public Function ShowTaskDialog( _
                 ByVal inContent As String, _
        Optional ByVal inButtons As TaskDialogButtons = TaskDialogButtons.tdOKCancel, _
        Optional ByVal inIcon As TaskDialogIcons = TaskDialogIcons.tdInformation, _
        Optional ByVal inTitle As String, _
        Optional ByVal inHeader As String, _
        Optional ByVal inParentHwnd As LongPtr = 0 _
    ) As VBA.VbMsgBoxResult
Rem .DESCRIPTION
Rem TaskDialog を表示する。
Rem .PARAMETER inContent
Rem ダイアログ内に表示するメッセージ。
Rem .PARAMETER inButtons
Rem 表示するボタン。省略時は「OK」ボタンと「キャンセル」ボタンが表示される。
Rem .PARAMETER inIcon
Rem タイトルとダイアログ内に表示するアイコン。省略時は青の○にiのアイコン。
Rem .PARAMETER inTitle
Rem ダイアログのタイトルに表示される文字列。省略時はホストアプリケーションのファイル名(Excelであれば"EXCEL.EXE")。
Rem .PARAMETER inHeader
Rem ダイアログ内で青字で表示される文字列。省略時は表示されない。
Rem .PARAMETER inParentHwnd
Rem 親とするウィンドウのハンドル。省略時はこのプロセスの最前面のウィンドウが親となる。
Rem .OUTPUTS
Rem ユーザーが押したボタン。
    
    '親ウィンドウが指定されなかった場合、このプロセスの最前面のウィンドウを親とする。
    If inParentHwnd = 0 Then inParentHwnd = GetActiveWindow()
    
    Const NullHInstance = 0
    Dim hr As HRESUT
    hr = TaskDialog( _
            inParentHwnd, _
            NullHInstance, _
            VBA.[_HiddenModule].StrPtr(inTitle), _
            VBA.[_HiddenModule].StrPtr(inHeader), _
            VBA.[_HiddenModule].StrPtr(inContent), _
            inButtons, _
            inIcon, _
            ShowTaskDialog _
        )
    Debug.Assert hr = S_OK
End Function
