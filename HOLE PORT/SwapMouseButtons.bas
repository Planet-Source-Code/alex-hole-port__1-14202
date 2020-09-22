Attribute VB_Name = "SwapMouseButtons"
Public Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long

'This swaps the left and right button of the mouse
Public Sub SwapButtons()
Dim Cur&, Butt&
    Cur = SwapMouseButton(Butt)
    If Cur = 0 Then
        SwapMouseButton (1)
    Else: SwapMouseButton (0)
    End If
End Sub
