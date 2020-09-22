Attribute VB_Name = "CaptureScreen"


Option Explicit
Option Base 0

Private Type PALETTEENTRY
   peRed As Byte
   peGreen As Byte
   peBlue As Byte
   peFlags As Byte
End Type

Private Type LOGPALETTE
   palVersion As Integer
   palNumEntries As Integer
   palPalEntry(255) As PALETTEENTRY
End Type

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Type PicBmp
   Size As Long
   Type As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long





'

Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
  Dim r As Long

   Dim Pic As PicBmp
  
   Dim IPic As IPicture
   Dim IID_IDispatch As GUID

On Error GoTo err

   With IID_IDispatch
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With

   With Pic
      .Size = Len(Pic)
      .Type = vbPicTypeBitmap
      .hBmp = hBmp
      .hPal = hPal
   End With


   r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

 
   Set CreateBitmapPicture = IPic
err:
Exit Function
End Function



  Public Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture

  Dim hDCMemory As Long
  Dim hBmp As Long
  Dim hBmpPrev As Long
  Dim r As Long
  Dim hDCSrc As Long
  Dim hPal As Long
  Dim hPalPrev As Long
  Dim RasterCapsScrn As Long
  Dim HasPaletteScrn As Long
  Dim PaletteSizeScrn As Long
  Dim LogPal As LOGPALETTE

 On Error GoTo err
   If Client Then
      hDCSrc = GetDC(hWndSrc)
   Else
      hDCSrc = GetWindowDC(hWndSrc)
   End If

   hDCMemory = CreateCompatibleDC(hDCSrc)
   
   hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
   hBmpPrev = SelectObject(hDCMemory, hBmp)


   RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)
                                                      
   HasPaletteScrn = RasterCapsScrn And RC_PALETTE
                                                        
   PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)
                                                        
   
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
     
      LogPal.palVersion = &H300
      LogPal.palNumEntries = 256
      r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
      hPal = CreatePalette(LogPal)
     
      hPalPrev = SelectPalette(hDCMemory, hPal, 0)
      r = RealizePalette(hDCMemory)
   End If

   
   r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)


   hBmp = SelectObject(hDCMemory, hBmpPrev)

   
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      hPal = SelectPalette(hDCMemory, hPalPrev, 0)
   End If

   
   r = DeleteDC(hDCMemory)
   r = ReleaseDC(hWndSrc, hDCSrc)

  
   Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
err:
Exit Function
End Function


'
Public Function CaptureScreen() As Picture
  Dim hWndScreen As Long

On Error GoTo err
   
   hWndScreen = GetDesktopWindow()

  

   Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
err:
Exit Function
End Function


Public Function CaptureForm(frmSrc As Form) As Picture
On Error GoTo err
   Set CaptureForm = CaptureWindow(frmSrc.hwnd, False, 0, 0, frmSrc.ScaleX(frmSrc.Width, vbTwips, vbPixels), frmSrc.ScaleY(frmSrc.Height, vbTwips, vbPixels))
err:
Exit Function
End Function


Public Function CaptureClient(frmSrc As Form) As Picture
On Error GoTo err
   Set CaptureClient = CaptureWindow(frmSrc.hwnd, True, 0, 0, frmSrc.ScaleX(frmSrc.ScaleWidth, frmSrc.ScaleMode, vbPixels), frmSrc.ScaleY(frmSrc.ScaleHeight, frmSrc.ScaleMode, vbPixels))
err:
Exit Function
End Function


Public Function CaptureActiveWindow() As Picture

    Dim hWndActive As Long
    Dim r As Long
    
    Dim RectActive As RECT
    
On Error GoTo err
    hWndActive = GetForegroundWindow()
    
   
    r = GetWindowRect(hWndActive, RectActive)
    
   
    Set CaptureActiveWindow = CaptureWindow(hWndActive, False, 0, 0, RectActive.Right - RectActive.Left, RectActive.Bottom - RectActive.Top)
err:
Exit Function
End Function


Public Sub PrintPictureToFitPage(Prn As Printer, Pic As Picture)
    
    Const vbHiMetric As Integer = 8
    Dim PicRatio As Double
    Dim PrnWidth As Double
    Dim PrnHeight As Double
    Dim PrnRatio As Double
    Dim PrnPicWidth As Double
    Dim PrnPicHeight As Double
On Error GoTo err
    If Pic.Height >= Pic.Width Then
        Prn.Orientation = vbPRORPortrait
    Else
        Prn.Orientation = vbPRORLandscape
    End If
    
   
    PicRatio = Pic.Width / Pic.Height
    
   
    PrnWidth = Prn.ScaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
    PrnHeight = Prn.ScaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)
   
    PrnRatio = PrnWidth / PrnHeight
    
    
    If PicRatio >= PrnRatio Then
       
        PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
        PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, Prn.ScaleMode)
    Else
       
        PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
        PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, Prn.ScaleMode)
    End If
    
    
    Prn.PaintPicture Pic, 0, 0, PrnPicWidth, PrnPicHeight

err:
Exit Sub
End Sub
'--------------
