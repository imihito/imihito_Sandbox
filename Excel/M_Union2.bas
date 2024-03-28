Attribute VB_Name = "M_Union2"
Option Explicit

Sub Sample()
'https://twitter.com/aero_iki/status/1356166832206450691
    
    Dim col As VBA.Collection 'Of Excel.Range
    Set col = New VBA.Collection
    
    Dim basRng As Excel.Range
    Set basRng = ActiveWindow.VisibleRange
    
    Dim cnt As Variant
    cnt = basRng.CountLarge
    
    Call VBA.Math.Randomize
    Dim i As Long
    For i = 1 To (cnt / 2)
        col.Add basRng.Item(cnt * VBA.Math.Rnd())
    Next i
    
    Dim t As Single
    Dim r As Excel.Range
    
    basRng.Clear
    t = Timer
    For Each r In col
        r.Interior.Color = vbRed
    Next r
    Debug.Print "Each", Format(Timer - t, "0.000")
    Stop
    
    basRng.Clear
    Dim xx As Excel.Range
    t = Timer
    For Each r In col
        If xx Is Nothing Then
            Set xx = r
        Else
            Set xx = xx.Application.Union(xx, r)
        End If
    Next r
    xx.Interior.Color = vbRed
    Debug.Print "Union", Format(Timer - t, "0.000")
    Stop
    
    
    basRng.Clear
    t = Timer
    Set xx = Union2(col)
    xx.Interior.Color = vbRed
    Debug.Print "Union2", Format(Timer - t, "0.000")
    Stop
    basRng.Clear
End Sub

Function Union2(ByVal inRanges As VBA.Collection) As Excel.Range 'Or Nothing
240329
    If inRanges.Count() = 0 Then
        Set Union2 = Nothing
        Exit Function
    End If
    
    Dim inputCnt As Long
    inputCnt = inRanges.Count()
    
    Dim rngs() As Excel.Range
    ReDim rngs(1 To inputCnt)
    
    Dim i As Long
    i = LBound(rngs) - 1
    
    Dim r As Excel.Range
    For Each r In inRanges
        i = i + 1
        Set rngs(i) = r
    Next r
    
    i = LBound(rngs)
    Dim rsltRng As Excel.Range
    Set rsltRng = rngs(i)
    i = i + 1
    
    Do While i < inputCnt
        Select Case i
            Case Is <= inputCnt - 28: Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6), rngs(i + 7), rngs(i + 8), rngs(i + 9), rngs(i + 10), rngs(i + 11), rngs(i + 12), rngs(i + 13), rngs(i + 14), rngs(i + 15), rngs(i + 16), rngs(i + 17), rngs(i + 18), rngs(i + 19), rngs(i + 20), rngs(i + 21), rngs(i + 22), rngs(i + 23), rngs(i + 24), rngs(i + 25), rngs(i + 26), rngs(i + 27), rngs(i + 28))
            Case inputCnt:      Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i))
            Case inputCnt - 1:  Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1))
            Case inputCnt - 2:  Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2))
            Case inputCnt - 3:  Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3))
            Case inputCnt - 4:  Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4))
            Case inputCnt - 5:  Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5))
            Case inputCnt - 6:  Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6))
            Case inputCnt - 7:  Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6), rngs(i + 7))
            Case inputCnt - 8:  Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6), rngs(i + 7), rngs(i + 8))
            Case inputCnt - 9:  Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6), rngs(i + 7), rngs(i + 8), rngs(i + 9))
            Case inputCnt - 10: Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6), rngs(i + 7), rngs(i + 8), rngs(i + 9), rngs(i + 10))
            Case inputCnt - 11: Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6), rngs(i + 7), rngs(i + 8), rngs(i + 9), rngs(i + 10), rngs(i + 11))
            Case inputCnt - 12: Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6), rngs(i + 7), rngs(i + 8), rngs(i + 9), rngs(i + 10), rngs(i + 11), rngs(i + 12))
            Case inputCnt - 13: Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6), rngs(i + 7), rngs(i + 8), rngs(i + 9), rngs(i + 10), rngs(i + 11), rngs(i + 12), rngs(i + 13))
            Case inputCnt - 14: Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6), rngs(i + 7), rngs(i + 8), rngs(i + 9), rngs(i + 10), rngs(i + 11), rngs(i + 12), rngs(i + 13), rngs(i + 14))
            Case inputCnt - 15: Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6), rngs(i + 7), rngs(i + 8), rngs(i + 9), rngs(i + 10), rngs(i + 11), rngs(i + 12), rngs(i + 13), rngs(i + 14), rngs(i + 15))
            Case inputCnt - 16: Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6), rngs(i + 7), rngs(i + 8), rngs(i + 9), rngs(i + 10), rngs(i + 11), rngs(i + 12), rngs(i + 13), rngs(i + 14), rngs(i + 15), rngs(i + 16))
            Case inputCnt - 17: Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6), rngs(i + 7), rngs(i + 8), rngs(i + 9), rngs(i + 10), rngs(i + 11), rngs(i + 12), rngs(i + 13), rngs(i + 14), rngs(i + 15), rngs(i + 16), rngs(i + 17))
            Case inputCnt - 18: Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6), rngs(i + 7), rngs(i + 8), rngs(i + 9), rngs(i + 10), rngs(i + 11), rngs(i + 12), rngs(i + 13), rngs(i + 14), rngs(i + 15), rngs(i + 16), rngs(i + 17), rngs(i + 18))
            Case inputCnt - 19: Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6), rngs(i + 7), rngs(i + 8), rngs(i + 9), rngs(i + 10), rngs(i + 11), rngs(i + 12), rngs(i + 13), rngs(i + 14), rngs(i + 15), rngs(i + 16), rngs(i + 17), rngs(i + 18), rngs(i + 19))
            Case inputCnt - 20: Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6), rngs(i + 7), rngs(i + 8), rngs(i + 9), rngs(i + 10), rngs(i + 11), rngs(i + 12), rngs(i + 13), rngs(i + 14), rngs(i + 15), rngs(i + 16), rngs(i + 17), rngs(i + 18), rngs(i + 19), rngs(i + 20))
            Case inputCnt - 21: Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6), rngs(i + 7), rngs(i + 8), rngs(i + 9), rngs(i + 10), rngs(i + 11), rngs(i + 12), rngs(i + 13), rngs(i + 14), rngs(i + 15), rngs(i + 16), rngs(i + 17), rngs(i + 18), rngs(i + 19), rngs(i + 20), rngs(i + 21))
            Case inputCnt - 22: Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6), rngs(i + 7), rngs(i + 8), rngs(i + 9), rngs(i + 10), rngs(i + 11), rngs(i + 12), rngs(i + 13), rngs(i + 14), rngs(i + 15), rngs(i + 16), rngs(i + 17), rngs(i + 18), rngs(i + 19), rngs(i + 20), rngs(i + 21), rngs(i + 22))
            Case inputCnt - 23: Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6), rngs(i + 7), rngs(i + 8), rngs(i + 9), rngs(i + 10), rngs(i + 11), rngs(i + 12), rngs(i + 13), rngs(i + 14), rngs(i + 15), rngs(i + 16), rngs(i + 17), rngs(i + 18), rngs(i + 19), rngs(i + 20), rngs(i + 21), rngs(i + 22), rngs(i + 23))
            Case inputCnt - 24: Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6), rngs(i + 7), rngs(i + 8), rngs(i + 9), rngs(i + 10), rngs(i + 11), rngs(i + 12), rngs(i + 13), rngs(i + 14), rngs(i + 15), rngs(i + 16), rngs(i + 17), rngs(i + 18), rngs(i + 19), rngs(i + 20), rngs(i + 21), rngs(i + 22), rngs(i + 23), rngs(i + 24))
            Case inputCnt - 25: Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6), rngs(i + 7), rngs(i + 8), rngs(i + 9), rngs(i + 10), rngs(i + 11), rngs(i + 12), rngs(i + 13), rngs(i + 14), rngs(i + 15), rngs(i + 16), rngs(i + 17), rngs(i + 18), rngs(i + 19), rngs(i + 20), rngs(i + 21), rngs(i + 22), rngs(i + 23), rngs(i + 24), rngs(i + 25))
            Case inputCnt - 26: Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6), rngs(i + 7), rngs(i + 8), rngs(i + 9), rngs(i + 10), rngs(i + 11), rngs(i + 12), rngs(i + 13), rngs(i + 14), rngs(i + 15), rngs(i + 16), rngs(i + 17), rngs(i + 18), rngs(i + 19), rngs(i + 20), rngs(i + 21), rngs(i + 22), rngs(i + 23), rngs(i + 24), rngs(i + 25), rngs(i + 26))
            Case inputCnt - 27: Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1), rngs(i + 2), rngs(i + 3), rngs(i + 4), rngs(i + 5), rngs(i + 6), rngs(i + 7), rngs(i + 8), rngs(i + 9), rngs(i + 10), rngs(i + 11), rngs(i + 12), rngs(i + 13), rngs(i + 14), rngs(i + 15), rngs(i + 16), rngs(i + 17), rngs(i + 18), rngs(i + 19), rngs(i + 20), rngs(i + 21), rngs(i + 22), rngs(i + 23), rngs(i + 24), rngs(i + 25), rngs(i + 26), rngs(i + 27))
        End Select
        i = i + 29
    Loop
    
    Set Union2 = rsltRng
End Function

Private Sub GenerateUnion2SelectCase()
    Dim i As Long
    For i = 1 To 28
        'Case inputCnt - 1:  Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i), rngs(i + 1))
        Debug.Print "Case inputCnt - "; i; ":Set rsltRng = rsltRng.Application.Union(rsltRng, rngs(i)";
        Dim j As Long
        For j = 1 To i
            Debug.Print ",rngs(i + "; j; ")";
        Next j
        Debug.Print ")"
    Next i
End Sub
