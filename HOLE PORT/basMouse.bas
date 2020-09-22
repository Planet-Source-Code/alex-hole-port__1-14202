Attribute VB_Name = "basMouse"
Option Explicit

'  Mouse/cursor functions.

Public lShowCursor As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

'
'  Hides the mouse cursor.
'
Public Sub HideMouse()
   Dim result As Integer
   
On Error GoTo err
   Do
      lShowCursor = lShowCursor - 1
      result = ShowCursor(False)
   Loop Until result < 0
err:
Exit Sub
End Sub


'
Public Sub RestoreMouse()
On Error GoTo err
   If lShowCursor > 0 Then
      Do While lShowCursor <> 0
         ShowCursor (False)
         lShowCursor = lShowCursor - 1
      Loop
   ElseIf lShowCursor < 0 Then
      Do While lShowCursor <> 0
         ShowCursor (True)
         lShowCursor = lShowCursor + 1
      Loop
   End If
err:
Exit Sub
End Sub


'
'  Show's the mouse cursor.
'
Public Sub ShowMouse()
   Dim result
On Error GoTo err
   Do
      lShowCursor = lShowCursor - 1
      result = ShowCursor(True)
   Loop Until result >= 0
err:
Exit Sub
End Sub

