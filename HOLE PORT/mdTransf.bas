Attribute VB_Name = "mdTransf"
Public flSize As Long
Public SavePath As String
Public NodeFullPath As String
Public NodeFullKey As String


Public Function EvalData(sIncoming As String, iRtLt As Integer, _
                  Optional sDivider As String) As String
   Dim i As Integer
   Dim tempStr As String
   ' Storage for the current Divider
   Dim sSplit As String
   
   ' the current character used to divide the data
   If sDivider = "" Then
      sSplit = ","
   Else
      sSplit = sDivider
   End If
   
   ' getting the right or left?
   Select Case iRtLt
        
      Case 1
          ' remove the data to the Left of the Current Divider
          For i = 0 To Len(sIncoming)
            tempStr = Left(sIncoming, i)
            
            If Right(tempStr, 1) = sSplit Then
              EvalData = Left(tempStr, Len(tempStr) - 1)
              Exit Function
            End If
          Next
          
      Case 2
          ' remove the data to the Right of the Current Divider
          For i = 0 To Len(sIncoming)
            tempStr = Right(sIncoming, i)
            
            If Left(tempStr, 1) = sSplit Then
              EvalData = Right(tempStr, Len(tempStr) - 1)
              Exit Function
            End If
          Next
   End Select
   
End Function

