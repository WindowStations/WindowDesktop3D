VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScreenShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'**************************************
'Windows API/Global Declarations for :Display Current Mouse Pointer Image
'**************************************
' Get the handle of the window the mouse is over
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
' Retrieves the handle of the current cursor
Private Declare Function GetCursor Lib "user32" () As Long
' Gets the coordinates of the mouse pointer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
' Gets the PID of the window specified
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
' Gets the PID of the current program
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
' This attaches our program to whichever thread "owns" the cursor at the moment
Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
' The next function draws the cursor to picCursor
' Note: If you want to display it in an Image control, use the GetDc API call
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
' The POINTAPI type hold the (X,Y) for GetCursorPos()
Private Type POINTAPI
   x As Long
   y As Long
End Type
' The following are used for keeping the window always on top. This is optional.
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_TOPMOST = -1
Private Const SWP_NOTOPMOST = -2
Private Const iconSize As Integer = 9
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function CreateCompatibleDC Lib "GDI32.DLL" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32.DLL" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32.DLL" (ByVal hdc As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "GDI32.DLL" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "GDI32.DLL" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectObject Lib "GDI32.DLL" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "GDI32.DLL" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal HDCSRC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "GDI32.DLL" (ByVal hdc As Long) As Long
Private Declare Function GetForegroundWindow Lib "USER32.DLL" () As Long
Private Declare Function SelectPalette Lib "GDI32.DLL" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "GDI32.DLL" (ByVal hdc As Long) As Long
Private Declare Function GetWindowDC Lib "USER32.DLL" (ByVal hWnd As Long) As Long
Private Declare Function GetDC Lib "USER32.DLL" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "USER32.DLL" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "USER32.DLL" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDesktopWindow Lib "USER32.DLL" () As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Type PicBmp
   Size As Long
Type As Long
   HBMP As Long
   hPal As Long
   Reserved As Long
End Type
Private Type PALETTEENTRY
   PERed As Byte
   PEGreen As Byte
   PEBlue As Byte
   PEFlags As Byte
End Type
Private Type LOGPALETTE
   PALVersion As Integer
   PALNumEntries As Integer
   PALPalEntry(255) As PALETTEENTRY
End Type
Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type
Private Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type
Private Const RASTERCAPS  As Long = 38
Private Const RC_PALETTE  As Long = &H100
Private Const SIZEPALETTE As Long = 104
Private Enum ImageFileFormat
   FF_RGB_BMP
   FF_RGB_PNG
   FF_RGB_JPG
   FF_ARGB_BMP
   FF_ARGB_PNG
End Enum
Private Type GdiplusStartupInput
   GdiplusVersion As Long
   DebugEventCallback As Long
   SuppressBackgroundThread As Long
   SuppressExternalCodecs As Long
End Type
Private Type BitmapData
   Width As Long
   Height As Long
   Stride As Long
   PixelFormat As Long
   Scan0 As Long
   Reserved As Long
End Type
Private Type RECTL
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type
Private Type Point
   x As Long
   y As Long
End Type
Private Type CLSID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type
Private Type EncoderParameter
   GUID As CLSID
   NumberOfValues As Long
Type As Long
   Value As Long
End Type
Private Type EncoderParameters
   Count As Long
   Parameter As EncoderParameter
End Type
Private Enum ImageLockMode
   ImageLockModeRead = &H1
   ImageLockModeWrite = &H2
   ImageLockModeUserInputBuf = &H4
End Enum
Private Const PixelFormat24bppRGB = &H21808
Private Const PixelFormat32bppARGB = &H26200A
Private Type BITMAPINFOHEADER
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Double
   biClrUsed As Double
End Type
Private Type BITMAPINFO
   bmiHeader As BITMAPINFOHEADER
   bmiColors As Long
End Type
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, InputBuf As GdiplusStartupInput, Optional ByVal OutputBuf As Long = 0) As Long
Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal pFileName As Long, BITMAP As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal BITMAP As Long, RECT As RECTL, ByVal flags As ImageLockMode, ByVal PixelFormat As Long, lockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal BITMAP As Long, lockedBitmapData As BitmapData) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal PixelFormat As Long, Scan0 As Any, BITMAP As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As Long, ByVal pFileName As Long, clsidEncoder As CLSID, encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As CLSID) As Long
' Paints the cursor image to the picturebox
Friend Function PaintCursor(PictureBox As PictureBox)
   Dim pt   As POINTAPI
   Dim hWnd As Long
   Dim dwThreadId, dwCurrentThreadID As Long
   Dim hCursor
   Dim threadid
   Dim CurrentThreadID
   ' Get the position of the cursor
   GetCursorPos pt
   ' Then get the handle of the window the cursor is over
   hWnd = WindowFromPoint(pt.x, pt.y)
   ' Get the PID of the thread
   threadid = GetWindowThreadProcessId(hWnd, vbNull)
   ' Get the thread of our program
   CurrentThreadID = App.threadid
   ' If the cursor is "owned" by a thread other than ours, attach to that thread and get the cursor
   If CurrentThreadID <> threadid Then
      AttachThreadInput CurrentThreadID, threadid, True
      hCursor = GetCursor()
      AttachThreadInput CurrentThreadID, threadid, False
      ' If the cursor is owned by our thread, use GetCursor() normally
   Else
      hCursor = GetCursor()
   End If
   ' Use DrawIcon to draw the cursor to picCursor
   DrawIcon PictureBox.hdc, pt.x - iconSize, pt.y - iconSize, hCursor
End Function
Friend Function CreateBitmapPicture(ByVal HBMP As Long, ByVal hPal As Long) As Picture
   Dim Pic           As PicBmp
   Dim IPic          As IPicture
   Dim IID_IDispatch As GUID
   On Error Resume Next
   With IID_IDispatch
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With
   With Pic
      .Size = Len(Pic)
      .Type = vbPicTypeBitmap
      .HBMP = HBMP
      .hPal = hPal
   End With
   OleCreatePictureIndirect Pic, IID_IDispatch, 1, IPic
   Set CreateBitmapPicture = IPic
End Function
Friend Function CaptureWindow(ByVal HWNDSrc As Long, ByVal Client As Boolean, ByVal LeftSRC As Long, ByVal TopSRC As Long, ByVal WidthSRC As Long, ByVal HeightSRC As Long) As Picture
   Dim HDCMemory       As Long
   Dim HBMP            As Long
   Dim HBMPPrev        As Long
   Dim HDCSRC          As Long
   Dim hPal            As Long
   Dim HPalPrev        As Long
   Dim RasterCapsScrn  As Long
   Dim HasPaletteScrn  As Long
   Dim PaletteSizeScrn As Long
   Dim LogPal          As LOGPALETTE
   On Error Resume Next
   If Client Then
      HDCSRC = GetDC(HWNDSrc)
   Else
      HDCSRC = GetWindowDC(HWNDSrc)
   End If
   HDCMemory = CreateCompatibleDC(HDCSRC)
   HBMP = CreateCompatibleBitmap(HDCSRC, WidthSRC, HeightSRC)
   HBMPPrev = SelectObject(HDCMemory, HBMP)
   RasterCapsScrn = GetDeviceCaps(HDCSRC, RASTERCAPS)
   HasPaletteScrn = RasterCapsScrn And RC_PALETTE
   PaletteSizeScrn = GetDeviceCaps(HDCSRC, SIZEPALETTE)
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      LogPal.PALVersion = &H300
      LogPal.PALNumEntries = 256
      GetSystemPaletteEntries HDCSRC, 0, 256, LogPal.PALPalEntry(0)
      hPal = CreatePalette(LogPal)
      HPalPrev = SelectPalette(HDCMemory, hPal, 0)
      RealizePalette HDCMemory
   End If
   BitBlt HDCMemory, 0, 0, WidthSRC, HeightSRC, HDCSRC, LeftSRC, TopSRC, vbSrcCopy
   HBMP = SelectObject(HDCMemory, HBMPPrev)
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      hPal = SelectPalette(HDCMemory, HPalPrev, 0)
   End If
   DeleteDC HDCMemory
   ReleaseDC HWNDSrc, HDCSRC
   Set CaptureWindow = CreateBitmapPicture(HBMP, hPal)
End Function
Friend Function CaptureScreen() As Picture
   Dim HWNDScreen As Long
   On Error Resume Next
   HWNDScreen = GetDesktopWindow()
   Set CaptureScreen = CaptureWindow(HWNDScreen, False, 0, 0, Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
End Function
Friend Function TakeScreenShot(PictureBox As PictureBox, ByVal FileOutput As String)
   PictureBox.Picture = CaptureScreen
   SavePicture PictureBox.Picture, FileOutput
End Function
