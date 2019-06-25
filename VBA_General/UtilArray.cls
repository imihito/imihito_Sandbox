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

'https://docs.microsoft.com/ja-jp/previous-versions/windows/desktop/api/oleauto/nf-oleauto-safearrayallocdescriptor
Private Declare PtrSafe Function SafeArrayAllocDescriptor Lib "oleaut32" ( _
        ByVal cDims As Long, _
        ByRef ppsaOut() As Any _
    ) As Long

Public Function CreateEmptyArray(iElementType As VBA.VbVarType) As Variant
    Select Case iElementType
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
    Let EmptyByteArray = VBA.vbNullString
End Property

#If Win64 Then
Public Property Get EmptyLongLongArray() As LongLong()
    Call SafeArrayAllocDescriptor(1, EmptyLongLongArray)
End Property
#End If

'配列の次元数を取得する。初期化していない動的配列の場合は0
Public Function DimensionsOf(ByRef iArray As Variant) As Long
    If Not VBA.IsArray(iArray) Then Err.Raise 13

    'VBAの多次元配列の次元の上限
    Const MAX_DIMENSION = 64

    Dim d As Long, no_mean_var As Long
    On Error Resume Next
        For d = 1 To MAX_DIMENSION
            no_mean_var = LBound(iArray, d)
            If Err.Number <> 0 Then Exit For
        Next d
    On Error GoTo 0

    'VBAのFor文は完走すると指定した数値+1になる
    Let DimensionsOf = (d - 1) '0 To 64
End Function

Public Function SizeOf( _
        ByRef iArray As Variant, _
        Optional iDimension As Long = 1 _
    ) As Long
    Let SizeOf = UBound(iArray, iDimension) - LBound(iArray, iDimension) + 1
End Function

Public Function CountAll(ByRef iArray As Variant) As Long
    Dim dimCnt As Long
    dimCnt = Me.DimensionsOf(iArray)
    If dimCnt = 0 Then Exit Function 'return 0
    
    Dim sumSize As Long
    sumSize = 1
    
    Dim d As Long
    For d = 1 To dimCnt
        sumSize = sumSize * Me.SizeOf(iArray, d)
    Next d
    
    Let CountAll = sumSize
End Function
