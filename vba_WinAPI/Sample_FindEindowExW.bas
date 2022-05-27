Attribute VB_Name = "Sample_FindEindowExW"
Option Explicit

'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-findwindowexw

Private Declare PtrSafe Function _
    FindWindowExW Lib "User32.dll" ( _
        ByVal hWndParent As LongPtr, _
        ByVal hWndChildAfter As LongPtr, _
        ByVal lpszClass As LongPtr, _
        ByVal lpszWindow As LongPtr _
    ) As LongPtr 'HWND

Private Const NullHWND As LongPtr = 0
Private Const NullStrPtr As LongPtr = 0
Private Const LongMaxValue = &H7FFFFFFF

Public Function FindWindows( _
        Optional ByVal inHwndParent As LongPtr = NullHWND, _
        Optional ByVal inClassName As String = VBA.Constants.vbNullString, _
        Optional ByVal inWindowText As String = VBA.Constants.vbNullString, _
        Optional ByVal inMaxDepth As Long = 1 _
    ) As VBA.Collection 'Of LongPtr(HWND)
    Const StartDepth = 1
    Set FindWindows = New VBA.Collection
    Call internalFindWindows( _
        inHwndParent, _
        VBA.[_HiddenModule].StrPtr(inClassName), _
        VBA.[_HiddenModule].StrPtr(inWindowText), _
        inMaxDepth, _
        StartDepth, _
        FindWindows _
    )
End Function

Private Sub internalFindWindows( _
        inHwndParent As LongPtr, _
        inClassName As LongPtr, _
        inWindowText As LongPtr, _
        inMaxDepth As Long, _
        inCurrentDepth As Long, _
        inFoundHwnds As VBA.Collection _
    )
    Dim foundHwnd As LongPtr
    foundHwnd = NullHWND
    
    Do  'まずは、条件に満たすウィンドウを探す。
        foundHwnd = FindWindowExW(inHwndParent, foundHwnd, inClassName, inWindowText)
        If foundHwnd = NullHWND Then Exit Do
        inFoundHwnds.Add foundHwnd
    Loop
    
    '再帰指定が無ければここで終了。
    If inCurrentDepth >= inMaxDepth Then Exit Sub
    
    Dim nextDepth As Long
    nextDepth = inCurrentDepth + 1
    
    foundHwnd = NullHWND 'なってるはずだけど念のため。
    
    Do  '条件指定なしで探索し、各ウィンドウについて子供でさらに検索。
        foundHwnd = FindWindowExW(inHwndParent, foundHwnd, NullStrPtr, NullStrPtr)
        If foundHwnd = NullHWND Then Exit Do
        
        Call internalFindWindows( _
                foundHwnd, _
                inClassName, _
                inWindowText, _
                inMaxDepth, _
                nextDepth, _
                inFoundHwnds _
            )
    Loop
    
End Sub
