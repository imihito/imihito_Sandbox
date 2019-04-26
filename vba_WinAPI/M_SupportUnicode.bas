Attribute VB_Name = "M_SupportUnicode"
Option Explicit

Private Declare PtrSafe Function MessageBoxW Lib "User32" ( _
        ByVal hWnd As LongPtr, _
        ByVal lpText As LongPtr, _
        ByVal lpCaption As LongPtr, _
        ByVal uType As VBA.VbMsgBoxStyle _
    ) As VBA.VbMsgBoxResult

Private Declare PtrSafe Function GetActiveWindow Lib "User32" ( _
    ) As LongPtr


Public Function MsgBox( _
        Prompt As String, _
        Optional Buttons As VBA.VbMsgBoxStyle = VBA.VbMsgBoxStyle.vbOKOnly, _
        Optional Title As String = "" _
    ) As VBA.VbMsgBoxResult
    
    Let MsgBox = MessageBoxW( _
            GetActiveWindow(), _
            VBA.StrPtr(Prompt), _
            VBA.StrPtr(Title), _
            Buttons _
        )
    
End Function

Private Sub Showponponpain()
    'VBA.MsgBox GetActiveWindow
    Const ponponpain = "&H70|&H36A|&H6F|&H363|&H6E|&H36C|&H70|&H363|&H6F|&H365|&H6E|&H36D|&H70|&H363|&H61|&H361|&H69|&H6E|&H365"
    
    Dim haraitai As String
    Dim tmp() As String
    tmp = VBA.Split(ponponpain, "|")
    Dim i As Long
    For i = LBound(tmp) To UBound(tmp)
        haraitai = haraitai & VBA.ChrW$(VBA.CLng(tmp(i)))
    Next i
    
    MsgBox haraitai, , haraitai
End Sub
