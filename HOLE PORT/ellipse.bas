Attribute VB_Name = "elipse"


Public Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" _
        (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
        ByVal Y2 As Long) As Long

Public Declare Function SetWindowRgn Lib "user32" _
        (ByVal hWnd As Long, ByVal hRgn As Long, _
        ByVal bRedraw As Boolean) As Long



Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Long) As Integer

Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
       Public Const SRCCOPY = &HCC0020
       Public Const SRCAND = &H8800C6
       Public Const SRCINVERT = &H660046

Declare Function GetDesktopWindow Lib "user32" () As Long
Public Dir1
Public LIN As Integer
Public NN As Integer

Public Sub Pause(ByVal nSecond As Single)
Dim t0 As Single
Dim dummy As Integer
        
        t0 = Timer
        
        Do While Timer - t0 < nSecond
                
                dummy = DoEvents()
                
                ' If we cross midnight, back up one day
                If Timer < t0 Then
                        t0 = t0 - 24 * 60 * 60 ' or t0 = t0 - 86400

                End If
        Loop

End Sub

