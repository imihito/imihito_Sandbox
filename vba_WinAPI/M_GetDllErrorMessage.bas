Attribute VB_Name = "M_GetDllErrorMessage"
Option Explicit

Private Declare PtrSafe Function _
    FormatMessage Lib "kernel32" Alias "FormatMessageA" ( _
        ByVal dwFlags As FORMAT_MESSAGE_FLAGS, _
        ByRef lpSource As Any, _
        ByVal dwMessageId As Long, _
        ByVal dwLanguageId As Long, _
        ByVal lpBuffer As String, _
        ByVal nSize As Long, _
        ByRef Arguments As LongPtr _
    ) As Long
    
Private Enum FORMAT_MESSAGE_FLAGS
    MAX_WIDTH_MASK = &HFF&
    
    ALLOCATE_BUFFER = &H100& 'FormatMessage 側で文字列領域を割り当ててもらう(結果の取得には要メモリ操作)。
    IGNORE_INSERTS = &H200&
    FROM_STRING = &H400&
    FROM_HMODULE = &H800&
    FROM_SYSTEM = &H1000& 'システムからメッセージを取得する(DLL関数のエラー取得時など)
    ARGUMENT_ARRAY = &H2000&
End Enum

'DLL 関数のエラーメッセージを取得する。
'dwMessageId    :エラーメッセージの Id。省略時は Err.LastDllError が使用される。
Public Function GetDllErrorMessage( _
        Optional ByVal dwMessageId As Long = 0 _
    ) As String
    
    '引数省略対応。
    If dwMessageId = 0 Then _
        dwMessageId = VBA.Information.Err().LastDllError
    
    'ALLOCATE_BUFFER を指定しないため、自前で領域を確保する。
    Dim paddingSize As Long
    paddingSize = &HFF
    Const paddingChar = VBA.Constants.vbNullChar
    
    Dim apiResult As Long
    Do
        'メッセージ用の領域確保。
        Dim lpBuffer As String
        lpBuffer = VBA.Strings.String$(paddingSize, paddingChar)
        Dim nSize As Long
        nSize = VBA.Strings.Len(lpBuffer)
        
        apiResult = FormatMessage( _
            FROM_SYSTEM Or MAX_WIDTH_MASK, _
            0, _
            dwMessageId, _
            0, _
            lpBuffer, _
            nSize, _
            0)
        
        '失敗時(≒領域不足時)は 0 になる。
        If apiResult <> 0 Then _
            Exit Do
        
        '確保サイズを大きくして再トライ。
        paddingSize = paddingSize * 2
    Loop
    
    '必要な範囲だけ取得して出力(apiResult の結果そのままは使いにくい)。
    Let GetDllErrorMessage = VBA.Strings.Left$(lpBuffer, VBA.Strings.InStr(1, lpBuffer, paddingChar) - 1)
End Function
