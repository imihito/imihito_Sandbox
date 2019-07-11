Attribute VB_Name = "TestMergeSort"
Option Explicit

Private Sub TestOfSortObject()
    Dim targetCol As VBA.Collection
    Set targetCol = New VBA.Collection
    
    Dim zeroCol As VBA.Collection
    Set zeroCol = New VBA.Collection
    
    Dim oneCol As VBA.Collection
    Set oneCol = New VBA.Collection
    oneCol.Add 1
    
    Dim twoCol As VBA.Collection
    Set twoCol = New VBA.Collection
    twoCol.Add 2
    twoCol.Add 2
    
    
    With targetCol
        .Add twoCol
        .Add zeroCol
        .Add oneCol
    End With
    
    Dim ascendCol As VBA.Collection
    Set ascendCol = SortObject(targetCol, "Count", VbMethod, True)
    
    With ascendCol
        Debug.Assert .Item(1) Is zeroCol
        Debug.Assert .Item(2) Is oneCol
        Debug.Assert .Item(3) Is twoCol
    End With
    
    
    Dim desecndCol As VBA.Collection
    Set desecndCol = SortObject(targetCol, "Count", VbMethod, False)
    
    With desecndCol
        Debug.Assert .Item(1) Is twoCol
        Debug.Assert .Item(2) Is oneCol
        Debug.Assert .Item(3) Is zeroCol
    End With
    
    Stop
    
End Sub


Private Sub TestOfSort1DimValueArray()
    Dim a() As Variant
    a = Array(1, "a", "A", "b")
    
    Dim b() As Variant
    b = Sort1DimValueArray(a, True, vbBinaryCompare)
    
    Debug.Assert "1Aab" = VBA.Strings.Join(b, VBA.Constants.vbNullString)
    
    Dim c() As Variant
    c = Sort1DimValueArray(a, True, vbTextCompare)
    
    Debug.Assert "1aAb" = VBA.Strings.Join(c, VBA.Constants.vbNullString)
    
    Dim d() As Variant
    d = Sort1DimValueArray(a, False, vbBinaryCompare)
    
    Debug.Assert "baA1" = VBA.Strings.Join(d, VBA.Constants.vbNullString)
    
    Stop
End Sub
