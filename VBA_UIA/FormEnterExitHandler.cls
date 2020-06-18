VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormEnterExitHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'UIAutomation の FocusChange イベントを使って、
'ユーザーフォームの Enter/Exit イベントを擬似的に再現するクラス。

'注意事項
    '- 参照設定として UIAutomationClient(C:\Windows\System32\UIAutomationCore.dll)が必須。
    '- オブジェクトのプロパティの「Instancing」を「2- PublicNotCreatable」に変更しておくこと。

'Events
    'FocusChanged   :監視対象のユーザーフォーム内でフォーカスの移動が起きたときに発生するイベント。
    
'Methods
    'Bind   :ユーザーフォームをこのクラスの監視対象として設定する。
    'Dispose:イベント購読を中止する。できるだけユーザーフォームを閉じる前に呼んでください。
    
Option Explicit

'監視対象のユーザーフォーム内でフォーカスの移動が起きたときに発生するイベント。
'CurrentFocusedControl  :現在フォーカスされているコントロール(=Enter)。
'PreviousFocusedControl :直前にフォーカスされていたコントロール(=Exit)。Nothing が渡されることがある。
Public Event FocusChanged( _
        ByVal CurrentFocusedControl As MSForms.Control, _
        ByVal PreviousFocusedControl As MSForms.Control _
    )

'UIAutomation 操作用の基本クラス。
Private cuia As UIAutomationClient.CUIAutomation

'FocusChangedEvent を受け取るためにインターフェイスを実装する。
Implements UIAutomationClient.IUIAutomationFocusChangedEventHandler


'直前/現在フォーカスされていたコントロール。
Private previousFocusedControl_ As MSForms.Control
Private currentFocusedControl_ As MSForms.Control


'監視対象の UI 要素と、MSForms.Control。
Private targetUIElements_() As UIAutomationClient.IUIAutomationElement
Private targetControls_() As MSForms.Control
    '備考
        'IUIAutomationElement を始め、UIAutomation 関連のほとんどのオブジェクトは IDispatch を実装していない。
        'そのため、遅延バインディングができず、型を明示しないとあらゆる操作ができない(DispCallFunc とかを使えば話は別)。
            'また Object 型変数に入らず、IsObject は False を返す。
        'この関係か、Variant 型変数にいれて型情報を消してしまうと様々な箇所で意図しない動作が起きやすい。
        'つまり、上記の配列は Scripting.Dictionary や VBA.Collection での代用は NG となる。


'監視対象の親のユーザーフォームの UI 要素。
Private formUiElement_ As UIAutomationClient.IUIAutomationElement
Private Const UserFormClassName = "F3 Server 05c80000"
'ElementFromIAccessible 経由の取得だと "ThunderDFrame" の一段下の上記のクラス名の要素となる。


'ユーザーフォームをこのクラスの監視対象として設定する。
Public Function Bind(ByVal inForm As MSForms.UserForm) As FormEnterExitHandler
    Set Bind = Me
    If inForm.Controls.Count = 0 Then Exit Function
    
    'フォーム内のコントールについて、UI 要素とコントロールそのものを取得する。
    ReDim targetUIElements_(0 To inForm.Controls.Count - 1)
    ReDim targetControls_(0 To inForm.Controls.Count - 1)
    
    Dim i As Long
    i = LBound(targetUIElements_)
    
    Dim c As MSForms.Control
    For Each c In inForm.Controls
        Set targetUIElements_(i) = cuia.ElementFromIAccessible(c, 0)
        Set targetControls_(i) = c
        i = i + 1
    Next c
    
    
    '親のフォームを保管する。
    Set formUiElement_ = cuia.ElementFromIAccessible(inForm, 0)
    Debug.Assert parentFormIsAlive()
    
    'イベント購読開始。
    Call cuia.AddFocusChangedEventHandler(cuia.CreateCacheRequest(), Me)
End Function


'イベント購読を中止する。
    '呼ばなくても自爆するようにはしているけれど、呼んだ方が良い。
Public Sub Dispose()
    cuia.RemoveAllEventHandlers
    
    Set previousFocusedControl_ = Nothing
    Set currentFocusedControl_ = Nothing
    Erase targetControls_, targetUIElements_
    targetControls_ = CVar(targetControls_)
End Sub

Private Sub Class_Initialize()
    targetControls_ = CVar(targetControls_) '0 To -1 の配列にする(targetUIElements_ に対してはこの方法では不可)。
    Set cuia = New UIAutomationClient.CUIAutomation
End Sub

Private Sub Class_Terminate()
    Call Me.Dispose 'あまり意味は無い(イベント購読中は CUIAutomation が参照を持っているため、Terminate されない)。
End Sub


'フォーカス変更イベントのハンドラー。
Private Sub IUIAutomationFocusChangedEventHandler_HandleFocusChangedEvent(ByVal sender As UIAutomationClient.IUIAutomationElement)
    If Not parentFormIsAlive() Then
        Call Me.Dispose
        Exit Sub
    End If
    
    Dim i As Long
    i = IndexOfUIElement(sender)
    If i < LBound(targetControls_) Then Exit Sub
    If targetControls_(i) Is currentFocusedControl_ Then Exit Sub
    
    Set previousFocusedControl_ = currentFocusedControl_
    Set currentFocusedControl_ = targetControls_(i)
    
    RaiseEvent FocusChanged(currentFocusedControl_, previousFocusedControl_)
End Sub


Private Function IndexOfUIElement(ByVal inUiaElement As UIAutomationClient.IUIAutomationElement) As Long
    Dim i As Long
    For i = LBound(targetControls_) To UBound(targetControls_) 'targetUIElements_ ではなく、targetControls_ なのはわざと(空配列対策)。
        If CBool(cuia.CompareElements(targetUIElements_(i), inUiaElement)) Then
            Let IndexOfUIElement = i
            Exit Function
        End If
    Next i
    Let IndexOfUIElement = LBound(targetControls_) - 1
End Function


Private Function parentFormIsAlive() As Boolean
    If formUiElement_ Is Nothing Then Exit Function
    Let parentFormIsAlive = (formUiElement_.CurrentClassName = UserFormClassName)
End Function
