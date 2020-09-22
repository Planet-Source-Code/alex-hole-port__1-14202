Attribute VB_Name = "PauseData"



Public Sub Pause(ByVal nSecond As Single)
Dim t0 As Single
Dim dummy As Integer
        
On Error GoTo err

        t0 = Timer
        
        Do While Timer - t0 < nSecond
                         dummy = DoEvents()
                

            
        Loop

err:
Exit Sub

End Sub
