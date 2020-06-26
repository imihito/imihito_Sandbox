VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Clipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NullHWND As LongPtr = 0
Private Const NullPtr As LongPtr = 0
Private Const CF_UNICODETEXT = 13

Private Enum BOOL
    bFALSE = 0
    bTRUE = 1
End Enum


'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-openclipboard
Private Declare PtrSafe Function _
    OpenClipboard Lib "User32.dll" ( _
        Optional ByVal hWndNewOwner As LongPtr = NullHWND _
    ) As BOOL

'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-closeclipboard
Private Declare PtrSafe Function _
    CloseClipboard Lib "User32.dll" ( _
    ) As BOOL


'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-emptyclipboard
Private Declare PtrSafe Function _
    EmptyClipboard Lib "User32.dll" ( _
    ) As BOOL


'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-isclipboardformatavailable
Private Declare PtrSafe Function _
    IsClipboardFormatAvailable Lib "User32.dll" ( _
        ByVal uFormat As Long _
    ) As BOOL


'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getclipboarddata
Private Declare PtrSafe Function _
    GetClipboardData Lib "User32.dll" ( _
        ByVal uFormat As Long _
    ) As LongPtr 'HANDLE


Private Declare PtrSafe Function _
    SetClipboardData Lib "User32.dll" ( _
        ByVal uFormat As Long, _
        ByVal hMem As LongPtr _
    ) As LongPtr 'HANDLE


Private Enum GMEM
    GMEM_FIXED = &H0&
    GMEM_MOVEABLE = &H2&
    GMEM_ZEROINIT = &H40&
    GPTR = GMEM_FIXED Or GMEM_ZEROINIT
    GHND = GMEM_MOVEABLE Or GMEM_ZEROINIT
End Enum

'https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-globalalloc
Private Declare PtrSafe Function _
    GlobalAlloc Lib "Kernel32.dll" ( _
        ByVal uFlags As GMEM, _
        ByVal dwBytes As Long _
    ) As LongPtr 'HGLOBAL

'https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-globalfree
Private Declare PtrSafe Function _
    GlobalFree Lib "Kernel32.dll" ( _
        ByVal hMem As LongPtr _
    ) As LongPtr 'HGLOBAL


'https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-globallock
Private Declare PtrSafe Function _
    GlobalLock Lib "Kernel32.dll" ( _
        ByVal hMem As LongPtr _
    ) As LongPtr 'LPVOID

'https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-globalunlock
Private Declare PtrSafe Function _
    GlobalUnlock Lib "Kernel32.dll" ( _
        ByVal hMem As LongPtr _
    ) As BOOL

'https://docs.microsoft.com/ja-jp/windows/win32/devnotes/rtlmovememory
Private Declare PtrSafe Sub _
    RtlMoveMemory Lib "Kernel32.dll" ( _
        ByVal Destination As LongPtr, _
        ByVal Source As LongPtr, _
        ByVal Length As Long _
    )


'https://docs.microsoft.com/ja-jp/windows/win32/api/oleauto/nf-oleauto-sysreallocstring
Private Declare PtrSafe Function _
    SysReAllocString Lib "OleAut32.dll" ( _
        ByVal pbstr As LongPtr, _
        ByVal psz As LongPtr _
    ) As BOOL

'以下は、エラー時の FormatMessageW 用各種定義。
Private Enum FORMAT_MESSAGE
    FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100&
    FORMAT_MESSAGE_IGNORE_INSERTS = &H200&
    FORMAT_MESSAGE_FROM_STRING = &H400&
    FORMAT_MESSAGE_FROM_HMODULE = &H800&
    FORMAT_MESSAGE_FROM_SYSTEM = &H1000&
    FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000&
    FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF&
End Enum

'https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-formatmessagew
Private Declare PtrSafe Function _
    FormatMessageW Lib "Kernel32.dll" ( _
        ByVal dwFlags As FORMAT_MESSAGE, _
        ByVal lpSource As LongPtr, _
        ByVal dwMessageId As Long, _
        ByVal dwLanguageId As Long, _
        ByVal lpBuffer As LongPtr, _
        ByVal nSize As Long, _
        ByVal Arguments As LongPtr _
    ) As Long

'https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-globalfree
Private Declare PtrSafe Function _
    LocalFree Lib "Kernel32.dll" ( _
        ByVal hMem As LongPtr _
    ) As LongPtr 'HLOCAL


'クリップボードから文字列を取得する。
Public Property Get Text() As String
    'クリップボードに文字列が含まれるか確認。
    If Not CBool(IsClipboardFormatAvailable(CF_UNICODETEXT)) Then Exit Property
    
    'クリップボードを開く。
    If Not CBool(OpenClipboard()) Then
        With VBA.Information.Err
            .Raise .LastDllError, , GetDllErrorMessage(.LastDllError)
        End With
    End If
    
    'エラー情報を格納する。
    Dim errStack As Variant
    errStack = getEmptyErrStack()
    
    'クリップボード上のデータ(文字列)へのハンドルを取得(=ポインタではない)。
    Dim hMem As LongPtr
    hMem = GetClipboardData(CF_UNICODETEXT)
    
    If hMem = NullPtr Then
        errStackPush errStack
        GoTo CLOSE_CLIP
    End If
    
    'ハンドルをロックして文字列へのポインタを取得。
    Dim bufferPtr As LongPtr
    bufferPtr = GlobalLock(hMem)
    
    If bufferPtr = NullPtr Then
        errStackPush errStack
        GoTo CLOSE_CLIP
    End If
    
    'ポインタから文字列を読み取り、Text(返り値)へ格納。
    If Not CBool(SysReAllocString(VBA.[_HiddenModule].VarPtr(Text), bufferPtr)) Then
        errStackPush errStack
        Call GlobalUnlock(hMem)
        GoTo CLOSE_CLIP
    End If
    
    'ロックを解放。
    Call GlobalUnlock(hMem)
    
CLOSE_CLIP:
    'クリップボードを閉じる。
    If Not CBool(CloseClipboard()) Then errStackPush errStack
    
    throwErrStack errStack
End Property


Public Property Let Text(ByVal inText As String)
    Dim srcStrPtr As LongPtr
    srcStrPtr = VBA.[_HiddenModule].StrPtr(inText)
    
    If srcStrPtr = NullPtr Then
        inText = ""
        srcStrPtr = VBA.[_HiddenModule].StrPtr(inText)
    End If
    
    Dim bufferSize As Long
    bufferSize = VBA.Strings.LenB(inText) & VBA.Strings.LenB(VBA.Constants.vbNullChar)
    
    Dim hMem As LongPtr
    hMem = allocGlobalData(srcStrPtr, bufferSize)
    
    Dim errStack As Variant
    errStack = getEmptyErrStack()
    
    'クリップボードを開く。
    If Not CBool(OpenClipboard()) Then
        errStackPush errStack
        If 0 <> GlobalFree(hMem) Then errStackPush errStack
        throwErrStack errStack
        Exit Property
    End If
    
    If Not CBool(EmptyClipboard()) Then _
        GoTo FREE_GLOBAL_AND_CLOSE_CLIP
    
    If 0 = SetClipboardData(CF_UNICODETEXT, hMem) Then _
        GoTo FREE_GLOBAL_AND_CLOSE_CLIP


CLOSE_CLIP:
    'クリップボードを閉じる。
    If Not CBool(CloseClipboard()) Then errStackPush errStack
    throwErrStack errStack
Exit Property

FREE_GLOBAL_AND_CLOSE_CLIP:
    errStackPush errStack
    If 0 <> GlobalFree(hMem) Then errStackPush errStack
    GoTo CLOSE_CLIP
End Property




Private Function allocGlobalData(inDataPtr As LongPtr, inLength As Long) As LongPtr
    Dim errStack As Variant
    errStack = getEmptyErrStack()
    
    Dim hMem As LongPtr
    hMem = GlobalAlloc(GMEM_MOVEABLE, inLength)
    
    If hMem = NullPtr Then GoTo FREE_GLOBAL
    
    Dim bufferPtr As LongPtr
    bufferPtr = GlobalLock(hMem)
    
    If bufferPtr = NullPtr Then GoTo FREE_GLOBAL 'ロックに失敗。
    
    RtlMoveMemory bufferPtr, inDataPtr, inLength
        
    'ロックを解放。
    If GlobalUnlock(hMem) <> 0 Then GoTo FREE_GLOBAL
    
    Let allocGlobalData = hMem
Exit Function

FREE_GLOBAL:
    errStackPush errStack
    If hMem <> NullPtr Then
        If 0 <> GlobalFree(hMem) Then errStackPush errStack
    End If
    throwErrStack errStack
End Function




Private Function getEmptyErrStack() As Variant
    Let getEmptyErrStack = VBA.[_HiddenModule].Array()
End Function

Private Sub errStackPush(inErrStack As Variant)
    ReDim Preserve inErrStack(LBound(inErrStack) To UBound(inErrStack) + 1)
    Let inErrStack(UBound(inErrStack)) = VBA.Information.Err().LastDllError
End Sub

Private Sub throwErrStack(inErrStack As Variant)
    Dim cnt As Long
    cnt = UBound(inErrStack) - LBound(inErrStack) + 1
    If cnt = 0 Then Exit Sub
    Dim i As Long
    For i = LBound(inErrStack) To UBound(inErrStack)
        inErrStack(i) = GetDllErrorMessage(inErrStack(i))
    Next i
    Err.Raise 5, , VBA.Strings.Join(inErrStack, VBA.Constants.vbNewLine)
End Sub



Private Function GetDllErrorMessage( _
        Optional ByVal inErrorNumber As Long = 0 _
    ) As String
    
    Dim msgId As Long
    If inErrorNumber <> 0 Then
        msgId = inErrorNumber
    Else
        msgId = VBA.Information.Err().LastDllError
    End If
    
    Dim bufPtr As LongPtr
    Dim n As Long
    n = FormatMessageW( _
        FORMAT_MESSAGE_ALLOCATE_BUFFER Or _
        FORMAT_MESSAGE_FROM_SYSTEM Or _
        FORMAT_MESSAGE_MAX_WIDTH_MASK, _
        0, _
        msgId, _
        0, _
        VBA.[_HiddenModule].VarPtr(bufPtr), _
        0, _
        0 _
    )
    If n = 0 Then
        Let GetDllErrorMessage = GetDllErrorMessage()
        Exit Function
    End If
    
    Debug.Assert CBool(SysReAllocString( _
        VBA.[_HiddenModule].VarPtr(GetDllErrorMessage), bufPtr _
    ))
    Debug.Assert 0 = LocalFree(bufPtr)
End Function

