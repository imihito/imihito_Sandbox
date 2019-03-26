Attribute VB_Name = "M_SendInput"
Option Explicit

Private Declare PtrSafe Function _
    SendInput Lib "User32.dll" ( _
        ByVal cInputs As Long, _
        ByVal pInputs As LongPtr, _
        ByVal cbSize As Long _
    ) As Long


Private Enum MouseEventFlag
    mfMOVE = &H1&          'mouse move
    mfLEFTDOWN = &H2&      'left button down
    mfLEFTUP = &H4&        'left button up
    mfRIGHTDOWN = &H8&     'right button down
    mfRIGHTUP = &H10&      'right button up
    mfMIDDLEDOWN = &H20&   'middle button down
    mfMIDDLEUP = &H40&     'middle button up
    mfXDOWN = &H80&
    mfXUP = &H100&          'An X button was released.
    mfWHEEL = &H800&
    mfHWHEEL = &H1000&
    mfMOVE_NOCOALESCE
    mfVIRTUALDESK = &H4000&
    mfABSOLUTE = &H8000&  'absolute move
End Enum

Private Type tagMOUSEINPUT
'https://docs.microsoft.com/en-us/windows/desktop/api/winuser/ns-winuser-tagmouseinput
    dx As Long
    dy As Long
    mouseData As Long
    dwFlags As Long
    timeStampMillisec As Long
    dwExtraInfo As LongPtr
End Type

Private Enum KeyboardEventFlag
    kfEXTENDEDKEY = &H1&
    kfKEYUP = &H2&
    kfSCANCODE = &H8&
    kfUNICODE = &H4&
End Enum


Private Type tagKEYBDINPUT
    'https://docs.microsoft.com/en-us/windows/desktop/api/winuser/ns-winuser-tagkeybdinput
    wVk As Integer
    wScan As Integer
    dwFlags As KeyboardEventFlag
    timeStampMillisec As Long
    dwExtraInfo As LongPtr
    
    padding8bit As Double '大きさを32バイト(tagMOUSEINPUT と同じ大きさ)にする。
End Type

Private Enum tiType
    tiINPUT_MOUSE = 0
    tiINPUT_KEYBOARD = 1
    tiINPUT_HARDWARE = 2
End Enum

Private Type tagINPUTkbd
    type As tiType ' = tiINPUT_KEYBOARD
    ki As tagKEYBDINPUT
End Type


Public Function SendText(iString As String) As Boolean
    Dim lenTxt As Long
    lenTxt = VBA.Len(iString)
    If lenTxt = 0 Then
        Let SendText = True
        Exit Function
    End If
    
    Dim inputs() As tagINPUTkbd
    ReDim inputs(1 To lenTxt * 2)
    
    Dim cInputs As Long
    cInputs = UBound(inputs) - LBound(inputs) + 1
    Dim pInputs As LongPtr
    pInputs = VBA.VarPtr(inputs(LBound(inputs)))
    Dim cbSize As Long
    cbSize = LenB(inputs(LBound(inputs)))
    
    Dim i As Long
    For i = LBound(inputs) To UBound(inputs) Step 2
        With inputs(i)
            .type = tiINPUT_KEYBOARD
            .ki.dwFlags = kfUNICODE
            .ki.wScan = VBA.AscW(VBA.Mid$(iString, (i + 1) / 2, 1))
        End With 'inputs(i)
        With inputs(i + 1)
            .type = tiINPUT_KEYBOARD
            .ki.dwFlags = kfKEYUP
        End With 'inputs(i + 1)
    Next i
    
    Dim resultOfSendInput As Long
    resultOfSendInput = SendInput(cInputs, pInputs, cbSize)
    
    Let SendText = (resultOfSendInput = cInputs)
End Function


Private Sub iaehf()
    'VBA.IsObject VBA.Time
    Dim mi As tagMOUSEINPUT
    Debug.Print "tagMOUSEINPUT", LenB(mi)
    Dim ki As tagKEYBDINPUT
    Debug.Print "tagKEYBDINPUT", LenB(ki)
    Dim a As Double
    Debug.Print LenB(a)
End Sub
