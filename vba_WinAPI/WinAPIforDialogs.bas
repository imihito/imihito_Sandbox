Attribute VB_Name = "WinAPIforDialogs"
Option Explicit

'https://memos-by-oxalis.hatenablog.com/entry/2019/09/22/223121
Public Enum TaskDialogButtons
    tdOk = 2 ^ 0
    tdYes = 2 ^ 1
    tdNo = 2 ^ 2
    tdCancel = 2 ^ 3
    tdRetry = 2 ^ 4
    tdClose = 2 ^ 5
End Enum

Public Enum TaskDialogIcons
    tdNone = 0&
    tdWarning = &HFFFF&
    tdError = &HFFFE&
    tdInformation = &HFFFD&
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
        ByVal pszIcon As TaskDialogIcons, _
        ByRef pnButton As VBA.VbMsgBoxResult _
    ) As HRESUT


Public Function ShowTaskDialog( _
                 ByVal inContent As String, _
        Optional ByVal inButtons As TaskDialogButtons = TaskDialogButtons.tdOk Or TaskDialogButtons.tdCancel, _
        Optional ByVal inIcon As TaskDialogIcons = TaskDialogIcons.tdInformation, _
        Optional ByVal inTitle As String, _
        Optional ByVal inHeader As String, _
        Optional ByVal inParentHwnd As LongPtr = 0 _
    ) As VBA.VbMsgBoxResult
    
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
