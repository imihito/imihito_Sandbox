VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'配列を要素数0で初期化するAPI。
    'cDimsは次元数。
    'ppsaOutは配列変数。
'https://docs.microsoft.com/ja-jp/previous-versions/windows/desktop/api/oleauto/nf-oleauto-safearrayallocdescriptor
Private Declare PtrSafe Function SafeArrayAllocDescriptor Lib "oleaut32" ( _
        ByVal cDims As Long, _
        ByRef ppsaOut() As Any _
    ) As Long

'要素数0の1次元配列を作成する。
Public Function CreateEmptyArray(inElementType As VBA.VbVarType) As Variant
    Select Case inElementType
        Case vbInteger:     Let CreateEmptyArray = EmptyIntegerArray
        Case vbLong:        Let CreateEmptyArray = EmptyLongArray
        Case vbSingle:      Let CreateEmptyArray = EmptySingleArray
        Case vbDouble:      Let CreateEmptyArray = EmptyDoubleArray
        Case vbCurrency:    Let CreateEmptyArray = EmptyCurrencyArray
        Case vbDate:        Let CreateEmptyArray = EmptyDateArray
        Case vbString:      Let CreateEmptyArray = EmptyStringArray
        Case vbObject:      Let CreateEmptyArray = EmptyObjectArray
        Case vbBoolean:     Let CreateEmptyArray = EmptyBooleanArray
        Case vbVariant:     Let CreateEmptyArray = EmptyVariantArray
        Case vbByte:        Let CreateEmptyArray = EmptyByteArray
    #If Win64 Then
        Case vbLongLong:    Let CreateEmptyArray = EmptyLongLongArray
    #End If
        Case Else: Err.Raise 13
    End Select
End Function

Public Property Get EmptyIntegerArray() As Integer()
    Call SafeArrayAllocDescriptor(1, EmptyIntegerArray)
End Property

Public Property Get EmptyLongArray() As Long()
    Call SafeArrayAllocDescriptor(1, EmptyLongArray)
End Property

Public Property Get EmptySingleArray() As Single()
    Call SafeArrayAllocDescriptor(1, EmptySingleArray)
End Property

Public Property Get EmptyDoubleArray() As Double()
    Call SafeArrayAllocDescriptor(1, EmptyDoubleArray)
End Property

Public Property Get EmptyCurrencyArray() As Currency()
    Call SafeArrayAllocDescriptor(1, EmptyCurrencyArray)
End Property

Public Property Get EmptyDateArray() As Date()
    Call SafeArrayAllocDescriptor(1, EmptyDateArray)
End Property

Public Property Get EmptyStringArray() As String()
    Let EmptyStringArray = VBA.Split(VBA.vbNullString)
End Property

Property Get EmptyObjectArray() As Variant '= Object()
    Dim o() As Object '-> Object()
    Let EmptyObjectArray = o
End Property '-> Object(0 To -1)

Public Property Get EmptyBooleanArray() As Boolean()
    Call SafeArrayAllocDescriptor(1, EmptyBooleanArray)
End Property

Public Property Get EmptyVariantArray() As Variant()
    Let EmptyVariantArray = VBA.Array()
End Property

Public Property Get EmptyByteArray() As Byte()
    Let EmptyByteArray = VBA.Constants.vbNullString
End Property

#If Win64 Then
Public Property Get EmptyLongLongArray() As LongLong()
    Call SafeArrayAllocDescriptor(1, EmptyLongLongArray)
End Property
#End If


'配列の次元数を取得する。初期化していない動的配列の場合は0。
Public Function DimensionsOf(ByRef inArray As Variant) As Long '0 To 64
    If Not VBA.Information.IsArray(inArray) Then Err.Raise 13
    
    'VBAの多次元配列の次元の上限。
    Const MAX_DIMENSION = 64

    Dim d As Long, no_mean_var As Long
    On Error Resume Next
        For d = 1 To MAX_DIMENSION
            no_mean_var = LBound(inArray, d)
            If Err.Number <> 0 Then Exit For
        Next d
    On Error GoTo 0

    'VBAのFor文は完走すると指定した数値+1になる。
    Let DimensionsOf = (d - 1) '0 To 64
End Function

'配列の要素数を取得する。
Public Function SizeOf( _
           ByRef inArray As Variant, _
        Optional inDimension As Long = 1 _
    ) As Long
    
    Let SizeOf = UBound(inArray, inDimension) - LBound(inArray, inDimension) + 1
End Function

Public Function CountAll(ByRef inArray As Variant) As Long
    Dim dimCnt As Long
    dimCnt = Me.DimensionsOf(inArray)
    If dimCnt = 0 Then Exit Function 'return 0
    
    Dim sumSize As Long
    sumSize = 1
    
    Dim d As Long
    For d = 1 To dimCnt
        sumSize = sumSize * Me.SizeOf(inArray, d)
    Next d
    
    Let CountAll = sumSize
End Function


'配列の下限も変えられる ReDim Preserve。
    'inoutOneDimArray
        '下限を変えたい1次元配列。
        'これを直接書き換える。
    'inLBound
        '新しい下限。
    'inUBound
        '新しい上限。
        '未指定時はinLBoundに合わせて自動調整。
        'inLengthと排他。inUBoundが優先。
    'inLength
        'ReDim後の配列の長さ。
        '未指定時はinoutOneDimArrayと同じ大きさ。
        
'参考:https://twitter.com/excelspeedup/status/1143362854856167424
Public Sub ReDimPreserve( _
           ByRef inoutOneDimArray As Variant, _
                 inLBound As Long, _
        Optional inUBound As Long = -2147483648#, _
        Optional inLength As Long = 0 _
    )
    Const LongMinValue As Long = -2147483648#
    
    '2次元配列以上をはじく。
    Dim dimCnt As Long
    dimCnt = Me.DimensionsOf(inoutOneDimArray)
    If dimCnt > 1 Then Err.Raise 13
    
    
    '以降は1次元配列か未初期化配列。
    
    Dim arraySize As Long
    If inUBound <> LongMinValue Then
        'inUBoundが指定されていれば大きさは求まる。
        arraySize = inUBound - inLBound + 1
    ElseIf inLength > 0 Then
        'inLengthが指定されている場合。
        arraySize = inLength
    Else
        'inUBound・inLengthともに未指定なら元の配列の大きさ(dimCnt = 0ならエラー)。
        arraySize = Me.SizeOf(inoutOneDimArray, 1)
    End If
    
    ReDim Preserve inoutOneDimArray(inLBound To inLBound + arraySize - 1)
    
End Sub

