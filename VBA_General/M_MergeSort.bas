Attribute VB_Name = "M_MergeSort"
Option Explicit

'オブジェクトのメンバーでソート(非破壊的処理)。
Public Function SortObject( _
                 inCollection As VBA.Collection, _
                 inProcName As String, _
        Optional inCallType As VBA.VbCallType = VBA.VbCallType.VbGet, _
        Optional inAscending As Boolean = True, _
        Optional inCompare As VBA.VbCompareMethod = VBA.VbCompareMethod.vbBinaryCompare _
    ) As VBA.Collection
    
    Select Case inCallType
        Case VBA.VbCallType.VbLet, _
             VBA.VbCallType.VbSet
            Call Err.Raise(5, "SortObject", "inCallTypeにはVbGetもしくはVbMethodを指定してください")
    End Select
    
    Dim sortKeys() As Variant
    ReDim sortKeys(1 To inCollection.Count())
    Dim objs() As Object
    ReDim objs(1 To inCollection.Count())
    
    Dim i As Long, o As Object
    i = LBound(sortKeys)
    For Each o In inCollection
        Let sortKeys(i) = VBA.Interaction.CallByName(o, inProcName, inCallType)
        Set objs(i) = o
        i = i + 1
    Next o
    
    'ソートした添え字を取得。
    Dim indexes() As Long
    indexes = getSortedIndexes(sortKeys, inAscending, inCompare)
    
    '出力用に入れ直し。
    Dim returnCol As VBA.Collection
    Set returnCol = New VBA.Collection
    For i = LBound(indexes) To UBound(indexes)
        returnCol.Add objs(indexes(i))
    Next i
    
    Set SortObject = returnCol
    
End Function


Public Function SortDictionaryByKey( _
                 inDictionary As Object, _
        Optional inAscending As Boolean = True _
    ) As Object 'As Scripting.Dictionary
    
    Dim sortKeys() As Variant
    sortKeys = inDictionary.Keys()
    
    Dim indexes() As Long
    indexes = getSortedIndexes(sortKeys, inAscending, inDictionary.CompareMode)
    
    Dim itms() As Variant
    itms = inDictionary.Items()
    
    Dim returnDic As Object 'As Scripting.Dictionary
    Set returnDic = VBA.Interaction.CreateObject("Scripting.Dictionary")
    returnDic.CompareMode = inDictionary.CompareMode
    
    Dim i As Long
    For i = LBound(indexes) To UBound(indexes)
        returnDic.Add sortKeys(indexes(i)), itms(indexes(i))
    Next i
    
    Set SortDictionaryByKey = returnDic
    
End Function



'マージソートの本体
    '値の一次元配列を渡すと、並びかえた場合の添え字の配列を返す。
    'inValues       :比較したい値の入った一次元配列(任意の型の配列を許容するため、Variant 型)
    'inAscending    :昇順か降順か。True なら昇順(小さい値から)。
    
    'return         :inValues の中身を inAscending の順番に取り出せる順番にした添え字配列

'e.g.
'getSortedIndexes(Array("a", "c", "b"), True) - > (0, 2, 1)
'getSortedIndexes(Array("b", "c", "b"), False) -> (1, 0, 2)
Private Function getSortedIndexes( _
        inValues As Variant, _
        inAscending As Boolean, _
        inCompare As VBA.VbCompareMethod _
    ) As Long()
    
    If Not VBA.Information.IsArray(inValues) Then Err.Raise 13
    
    Dim inUseUpperResult As Long
    If inAscending Then
        inUseUpperResult = 1
    Else
        inUseUpperResult = -1
    End If
    
    
    '`inValues`を元に添え字配列作成
    Dim basIndexes() As Long
    ReDim basIndexes(LBound(inValues) To UBound(inValues))
    Dim i As Long
    For i = LBound(inValues) To UBound(inValues)
        basIndexes(i) = i
    Next i
    
    '添え字配列をコピー。
    getSortedIndexes = basIndexes
    
    Call recurseMergeSort( _
            inValues, _
            basIndexes, _
            getSortedIndexes, _
            LBound(inValues), _
            UBound(inValues) - LBound(inValues) + 1, _
            inUseUpperResult, _
            inCompare _
        )
    
End Function


Private Sub recurseMergeSort( _
        inValues As Variant, _
        inSrcIndexes() As Long, _
        outDestIndexes() As Long, _
        inStart As Long, _
        inLength As Long, _
        inUseUpperResult As Long, _
        inCompare As VBA.VbCompareMethod _
    )

    Dim halfLen As Long
    halfLen = CLng(inLength / 2)

    '前半部分をソート
    If halfLen >= 2 Then _
        Call recurseMergeSort(inValues, outDestIndexes, inSrcIndexes, inStart, halfLen, inUseUpperResult, inCompare)
    '後半部分をソート
    If inLength - halfLen >= 2 Then _
        Call recurseMergeSort(inValues, outDestIndexes, inSrcIndexes, inStart + halfLen, inLength - halfLen, inUseUpperResult, inCompare)
    
    
    '前半部分の添え字と最大値
    Dim lwIndex As Long
    lwIndex = inStart
    Dim lwLimit As Long
    lwLimit = inStart + halfLen - 1
    
    '後半部分の添え字と最大値
    Dim upIndex As Long
    upIndex = inStart + halfLen
    Dim upLimit As Long
    upLimit = inStart + inLength - 1
    
    'ソート後配列の添え字と最大値
    Dim destIndex As Long
    destIndex = inStart
    Dim destLimit As Long
    destLimit = inStart + inLength - 1

    Dim remainIndex As Long '片方が終わった時の余り用
    
    For destIndex = inStart To destLimit Step 1
        If compareM( _
                inValues(inSrcIndexes(lwIndex)), _
                inValues(inSrcIndexes(upIndex)), _
                inCompare _
            ) = inUseUpperResult Then
            
            'upIndexの方が値が小さい場合（昇順の場合）
            outDestIndexes(destIndex) = inSrcIndexes(upIndex)
            
            If upIndex = upLimit Then
                remainIndex = lwIndex
                Exit For
            End If
            upIndex = upIndex + 1
            
        Else
            '値が同じ or lwIndexの方が値が小さい場合（昇順の場合）
            outDestIndexes(destIndex) = inSrcIndexes(lwIndex)
            
            If lwIndex = lwLimit Then
                remainIndex = upIndex
                Exit For
            End If
            lwIndex = lwIndex + 1
            
        End If
    Next destIndex
    
    '残りの分を代入
    'Next destIndexを飛ばした分インクリメントしてからスタート
    For destIndex = destIndex + 1 To destLimit Step 1
        outDestIndexes(destIndex) = inSrcIndexes(remainIndex)
        remainIndex = remainIndex + 1
    Next destIndex

End Sub

Private Function compareM( _
        inValue1 As Variant, _
        inValue2 As Variant, _
        inCompare As VBA.VbCompareMethod _
    ) As Long
    Select Case VBA.VbVarType.vbString
        Case VBA.Information.VarType(inValue1), _
             VBA.Information.VarType(inValue2)
            Let compareM = VBA.Strings.StrComp(inValue1, inValue2, inCompare)
            Exit Function
    End Select
    
    
    If inValue1 < inValue2 Then Let compareM = -1: Exit Function
    If inValue1 > inValue2 Then Let compareM = 1:  Exit Function
    Let compareM = 0
End Function

