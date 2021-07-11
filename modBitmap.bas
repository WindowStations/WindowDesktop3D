Attribute VB_Name = "modBitmap"
Option Explicit
'=============================================================================================================
'
' modBitmap Module
' ----------------
'
' Created By  : Kevin Wilson
'               http://www.TheVBZone.com   ( The VB Zone )
'               http://www.TheVBZone.net   ( The VB Zone .net )
'
' Created On  : January 17, 2001
' Last Update : October 05, 2003
'
' VB Versions : 5.0 / 6.0
'
' Requires    : NOTHING
'
' Description : This module is a collection of very useful functions designed for use with graphics
'               manipulation.  You can use it to find out information about pictures in memory, or to
'               render such pictures on to specified Device Contexts (DC).
'
'               If any of the functions in this module fail, you can most likely get extended error
'               information by calling the "GetLastError" Win32 API function.
'
' Example Use :
'
'  Dim TheHeight As Long
'  Dim TheWidth  As Long
'
'  ' Show the form so you can see the drawing as it happens in debug mode (step by step)
'  Me.Show
'
'  ' Make the picture redraw itself so it doesn't go away if the form loses focus
'  Picture1.AutoRedraw = True
'
'  ' Get the height/width in pixels
'  If Convert_HM_PX(Me.Picture.Height, Me.Picture.Width, TheHeight, TheWidth, True) = True Then
'
'    ' Draw the picture onto the PictureBox object called "Picture1" and
'    ' invert the colors by specifying "SCRINVERT" as the raster operation
'    ' instead of the default "SCRCOPY"
'    If RenderBitmapEx(Picture1.hDC, , Me.Picture.Handle, 0, 0, 0, 0, TheHeight, TheWidth, SRCINVERT, , , , True, Picture1.hWnd) = True Then
'      Debug.Print "SUCCESS!"
'    Else
'      Debug.Print "FAILED!"
'    End If
'  End If
'
'  ' Set the "Picture" property to equal the "Image" property
'  Set Picture1.Picture = Picture1.Image
'
'=============================================================================================================
'
' LEGAL:
'
' You are free to use this code as long as you keep the above heading information intact and unchanged. Credit
' given where credit is due.  Also, it is not required, but it would be appreciated if you would mention
' somewhere in your compiled program that that your program makes use of code written and distributed by
' Kevin Wilson (www.TheVBZone.com).  Feel free to link to this code via your web site or articles.
'
' You may NOT take this code and pass it off as your own.  You may NOT distribute this code on your own server
' or web site.  You may NOT take code created by Kevin Wilson (www.TheVBZone.com) and use it to create products,
' utilities, or applications that directly compete with products, utilities, and applications created by Kevin
' Wilson, TheVBZone.com, or Wilson Media.  You may NOT take this code and sell it for profit without first
' obtaining the written consent of the author Kevin Wilson.
'
' These conditions are subject to change at the discretion of the owner Kevin Wilson at any time without
' warning or notice.  Copyright© by Kevin Wilson.  All rights reserved.
'
'=============================================================================================================
' Type - General
Public Type RECT
   left   As Long
   top    As Long
   right  As Long
   bottom As Long
End Type
' Type - General
Public Type POINTAPI
   x As Long
   y As Long
End Type
' Type - GetEnhMetaFileHeader.lpEMH.(rclBounds/rclFrame)
Public Type RECTL
   left   As Long
   top    As Long
   right  As Long
   bottom As Long
End Type
' Type - GetEnhMetaFileHeader.lpEMH.(szlDevice/szlMillimeters/szlMicrometers)
Public Type SIZEL
   cx As Long
   cy As Long
End Type
' Type - OleCreatePictureIndirect
Public Type GUID
   Data1    As Long
   Data2    As Integer
   Data3    As Integer
   Data4(7) As Byte
End Type
' Type - OleCreatePictureIndirect / OleLoadPicture
Public Type PICTDESC_ALL
   cbSizeOfStruct As Long 'UINT     // Size of the PICTDESC structure.
   PicType        As Long 'UINT     // Type of picture described by this structure, which can be any of the following values: PICTYPE_UNINITIALIZED, PICTYPE_NONE, PICTYPE_BITMAP, PICTYPE_METAFILE, PICTYPE_ICON, PICTYPE_ENHMETAFILE
   hPicture       As Long 'LPVLOID  // Pointer to the bits that make up the picture.  This varies depending on the type of picture (see following structures)
   hPalette       As Long 'HPALETTE // Pointer to the picture's palette (where applicable)
   Reserved       As Long '         // Reserved
End Type
' Type - OleCreatePictureIndirect / OleLoadPicture
Public Type PICTDESC_BMP 'picType = PICTYPE_BITMAP
   cbSizeOfStruct As Long 'UINT     // Size of the PICTDESC structure.
   PicType        As Long 'UINT     // Type of picture described by this structure, which can be any of the following values: PICTYPE_UNINITIALIZED, PICTYPE_NONE, PICTYPE_BITMAP, PICTYPE_METAFILE, PICTYPE_ICON, PICTYPE_ENHMETAFILE
   hBitmap        As Long 'HBITMAP  // The HBITMAP identifying the bitmap assigned to the picture object.
   hPal           As Long 'HPALETTE // The HPALETTE identifying the color palette for the bitmap.
End Type
' Type - OleCreatePictureIndirect / OleLoadPicture
Public Type PICTDESC_META 'picType = PICTYPE_METAFILE
   cbSizeOfStruct As Long 'UINT      // Size of the PICTDESC structure.
   PicType        As Long 'UINT      // Type of picture described by this structure, which can be any of the following values: PICTYPE_UNINITIALIZED, PICTYPE_NONE, PICTYPE_BITMAP, PICTYPE_METAFILE, PICTYPE_ICON, PICTYPE_ENHMETAFILE
   hMeta          As Long 'HMETAFILE // The HMETAFILE handle identifying the metafile assigned to the picture object.
   xExt           As Long 'int       // Horizontal extent of the metafile in HIMETRIC units.
   yExt           As Long 'int       // Vertical extent of the metafile in HIMETRIC units.
End Type
' Type - OleCreatePictureIndirect / OleLoadPicture
Public Type PICTDESC_ICON 'picType = PICTYPE_ICON
   cbSizeOfStruct As Long 'UINT  // Size of the PICTDESC structure.
   PicType        As Long 'UINT  // Type of picture described by this structure, which can be any of the following values: PICTYPE_UNINITIALIZED, PICTYPE_NONE, PICTYPE_BITMAP, PICTYPE_METAFILE, PICTYPE_ICON, PICTYPE_ENHMETAFILE
   hIcon          As Long 'HICON // The HICON identifying the icon assigned to the picture object.
End Type
' Type - OleCreatePictureIndirect / OleLoadPicture
Public Type PICTDESC_EMETA 'picType = PICTYPE_ENHMETAFILE
   cbSizeOfStruct As Long 'UINT         // Size of the PICTDESC structure.
   PicType        As Long 'UINT         // Type of picture described by this structure, which can be any of the following values: PICTYPE_UNINITIALIZED, PICTYPE_NONE, PICTYPE_BITMAP, PICTYPE_METAFILE, PICTYPE_ICON, PICTYPE_ENHMETAFILE
   hEMF           As Long 'HENHMETAFILE // The HENHMETAFILE identifying the enhanced metafile to assign to the picture object.
End Type
' Type - GetObjectAPI.lpObject
Public Type BITMAP
   bmType       As Long    'LONG   // Specifies the bitmap type. This member must be zero.
   bmWidth      As Long    'LONG   // Specifies the width, in pixels, of the bitmap. The width must be greater than zero.
   bmHeight     As Long    'LONG   // Specifies the height, in pixels, of the bitmap. The height must be greater than zero.
   bmWidthBytes As Long    'LONG   // Specifies the number of bytes in each scan line. This value must be divisible by 2, because Windows assumes that the bit values of a bitmap form an array that is word aligned.
   bmPlanes     As Integer 'WORD   // Specifies the count of color planes.
   bmBitsPixel  As Integer 'WORD   // Specifies the number of bits required to indicate the color of a pixel.
   bmBits       As Long    'LPVOID // Points to the location of the bit values for the bitmap. The bmBits member must be a long pointer to an array of character (1-byte) values.
End Type
' Type - CreateIconIndirect / GetIconInfo
Public Type ICONINFO
   fIcon    As Long 'BOOL    // Specifies whether this structure defines an icon or a cursor. A value of TRUE specifies an icon; FALSE specifies a cursor.
   xHotspot As Long 'DWORD   // Specifies the x-coordinate of a cursor’s hot spot. If this structure defines an icon, the hot spot is always in the center of the icon, and this member is ignored.
   yHotspot As Long 'DWORD   // Specifies the y-coordinate of the cursor’s hot spot. If this structure defines an icon, the hot spot is always in the center of the icon, and this member is ignored.
   hbmMask  As Long 'HBITMAP // Specifies the icon bitmask bitmap. If this structure defines a black and white icon, this bitmask is formatted so that the upper half is the icon AND bitmask and the lower half is the icon XOR bitmask. Under this condition, the height should be an even multiple of two. If this structure defines a color icon, this mask only defines the AND bitmask of the icon.
   hbmColor As Long 'HBITMAP // Identifies the icon color bitmap. This member can be optional if this structure defines a black and white icon. The AND bitmask of hbmMask is applied with the SRCAND flag to the destination; subsequently, the color bitmap is applied (using XOR) to the destination by using the SRCINVERT flag.
End Type
' Type - GetEnhMetaFileHeader.lpEMH
Public Type ENHMETAHEADER
   iType          As Long    'DWORD // Specifies the record type. This member must specify the value assigned to the EMR_HEADER constant.
   nSize          As Long    'DWORD // Specifies the structure size, in bytes.
   rclBounds      As RECTL   'RECTL // Specifies the dimensions, in device units, of the smallest rectangle that can be drawn around the picture stored in the metafile. This rectangle is supplied by graphics device interface (GDI). Its dimensions include the right and bottom edges.
   rclFrame       As RECTL   'RECTL // Specifies the dimensions, in .01 millimeter units, of a rectangle that surrounds the picture stored in the metafile. This rectangle must be supplied by the application that creates the metafile. Its dimensions include the right and bottom edges.
   dSignature     As Long    'DWORD // Specifies a double word signature. This member must specify the value assigned to the ENHMETA_SIGNATURE constant.
   nVersion       As Long    'DWORD // Specifies the metafile version. The current version value is 0x10000.
   nBytes         As Long    'DWORD // Specifies the size of the enhanced metafile, in bytes.
   nRecords       As Long    'DWORD // Specifies the number of records in the enhanced metafile.
   nHandles       As Integer 'WORD  // Specifies the number of handles in the enhanced-metafile handle table. (Index zero in this table is reserved.)
   sReserved      As Integer 'WORD  // Reserved; must be zero.
   nDescription   As Long    'DWORD // Specifies the number of characters in the array that contains the description of the enhanced metafile's contents. This member should be set to zero if the enhanced metafile does not contain a description string.
   offDescription As Long    'DWORD // Specifies the offset from the beginning of the ENHMETAHEADER structure to the array that contains the description of the enhanced metafile's contents. This member should be set to zero if the enhanced metafile does not contain a description string.
   nPalEntries    As Long    'DWORD // Specifies the number of entries in the enhanced metafile's palette.
   szlDevice      As SIZEL   'SIZEL // Specifies the resolution of the reference device, in pixels.
   szlMillimeters As SIZEL   'SIZEL // Specifies the resolution of the reference device, in millimeters.
   cbPixelFormat  As Long    'DWORD // Windows 95/98, Windows NT4.0 and later: Specifies the size of the last recorded pixel format in a metafile. If a pixel format is set in a reference DC at the start of recording, cbPixelFormat is set to the size of the PIXELFORMATDESCRIPTOR. When no pixel format is set when a metafile is recorded, this member is set to zero. If more than a single pixel format is set, the header points to the last pixel format.
   offPixelFormat As Long    'DWORD // Windows 95/98, Windows NT4.0 and later: Specifies the offset of pixel format used when recording a metafile. If a pixel format is set in a reference DC at the start of recording or during recording, offPixelFormat is set to the offset of the PIXELFORMATDESCRIPTOR in the metafile. If no pixel format is set when a metafile is recorded, this member is set to zero. If more than a single pixel format is set, the header points to the last pixel format.
   bOpenGL        As Long    'DWORD // Windows 95/98, Windows NT4.0 and later: Specifies whether any OpenGL records are present in a metafile. bOpenGL is a simple Boolean flag that you can use to determine whether an enhanced metafile requires OpenGL handling. When a metafile contains OpenGL records, bOpenGL is TRUE; otherwise it is FALSE.
   ' szlMicrometers As SIZEL   'SIZEL // Windows 98,    Windows 2000           : Size of the reference device in micrometers.
End Type
' Constants - BitBlt.dwRop
Public Enum RasterOperations
   SRCCOPY = &HCC0020          ' Copies the source rectangle directly to the destination rectangle.
   SRCPAINT = &HEE0086         ' Combines the colors of the source and destination rectangles by using the Boolean OR operator.
   SRCAND = &H8800C6           ' Combines the colors of the source and destination rectangles by using the Boolean AND operator.
   SRCINVERT = &H660046        ' Combines the colors of the source and destination rectangles by using the Boolean XOR operator.
   SRCERASE = &H440328         ' Combines the inverted colors of the destination rectangle with the colors of the source rectangle by using the Boolean AND operator.
   NOTSRCCOPY = &H330008       ' Copies the inverted source rectangle to the destination.
   NOTSRCERASE = &H1100A6      ' Combines the colors of the source and destination rectangles by using the Boolean OR operator and then inverts the resultant color.
   MERGECOPY = &HC000CA        ' Merges the colors of the source rectangle with the brush currently selected in hdcDest, by using the Boolean AND operator.
   MERGEPAINT = &HBB0226       ' Merges the colors of the inverted source rectangle with the colors of the destination rectangle by using the Boolean OR operator.
   PATCOPY = &HF00021          ' Copies the brush currently selected in hdcDest, into the destination bitmap.
   PATPAINT = &HFB0A09         ' Combines the colors of the brush currently selected in hdcDest, with the colors of the inverted source rectangle by using the Boolean OR operator. The result of this operation is combined with the colors of the destination rectangle by using the Boolean OR operator.
   PATINVERT = &H5A0049        ' Combines the colors of the brush currently selected in hdcDest, with the colors of the destination rectangle by using the Boolean XOR operator.
   DSTINVERT = &H550009        ' Inverts the destination rectangle.
   BLACKNESS = &H42            ' Fills the destination rectangle using the color associated with index 0 in the physical palette. (This color is black for the default physical palette.)
   WHITENESS = &HFF0062        ' Fills the destination rectangle using the color associated with index 1 in the physical palette. (This color is white for the default physical palette.)
   NOMIRRORBITMAP = &H80000000 ' Windows 98, Windows 2000: Prevents the bitmap from being mirrored.
   CAPTUREBLT = &H40000000     ' Windows 98, Windows 2000: Includes any windows that are layered on top of your window in the resulting image. By default, the image only contains your window.
End Enum
' Constants - LoadResData
Public Enum ResTypes
   RT_BITMAP = vbResBitmap
   RT_ICON = vbResIcon
   RT_CURSOR = vbResCursor
   rt_Custom = 3
End Enum
' Constants - BITMAP.bmType & CopyImage.uType
Public Enum PictureTypes
   IMAGE_BITMAP = 0
   IMAGE_CURSOR = 1
   IMAGE_ICON = 2
   IMAGE_ENHMETAFILE = 3
End Enum
' Constants - General
Public Const MAX_PATH = 260
' Constants - ENHMETAHEADER.iType
Public Const EMR_HEADER = 1
' Constants - ENHMETAHEADER.dSignature
Public Const ENHMETA_SIGNATURE = &H20454D46
' Constants - PICTDESC.picType
Public Const PICTYPE_UNINITIALIZED = -1 ' The picture object is currently uninitialized.
Public Const PICTYPE_NONE = 0           ' A new picture object is to be created without an initialized state. This value is valid only in the PICTDESC structure.
Public Const PICTYPE_BITMAP = 1         ' The picture type is a bitmap. When this value occurs in the PICTDESC structure, it means that the bmp field of that structure contains the relevant initialization parameters.
Public Const PICTYPE_METAFILE = 2       ' The picture type is a metafile. When this value occurs in the PICTDESC structure, it means that the wmf field of that structure contains the relevant initialization parameters.
Public Const PICTYPE_ICON = 3           ' The picture type is an icon. When this value occurs in the PICTDESC structure, it means that the icon field of that structure contains the relevant initialization parameters.
Public Const PICTYPE_ENHMETAFILE = 4    ' The picture type is a Win32-enhanced metafile. When this value occurs in the PICTDESC structure, it means that the emf field of that structure contains the relevant initialization parameters.
' Constants - GetDeviceCaps.nIndex
Public Const HORZSIZE = 4    ' Width, in millimeters, of the physical screen.
Public Const VERTSIZE = 6    ' Height, in millimeters, of the physical screen.
Public Const HORZRES = 8     ' Width, in pixels, of the screen.
Public Const VERTRES = 10    ' Height, in raster lines, of the screen.
Public Const BITSPIXEL = 12  ' Number of adjacent color bits for each pixel.
' Constants - OleCreateBitmapIndiect (Return Values)
Public Const S_OK = 0                   ' The new picture object was created successfully.
Public Const E_NOINTERFACE = &H80004002 ' The object does not support the interface specified in riid.
Public Const E_POINTER = &H80004003     ' The address in pPictDesc or ppvObj is not valid. For example, it may be NULL.
Public Const E_INVALIDARG = &H80000003  ' One or more arguments are invalid
Public Const E_OUTOFMEMORY = &H8007000E ' Ran out of memory
Public Const E_UNEXPECTED = &H8000FFFF  ' Catastrophic failure
' Constants - GetCurrentObject.uObjectType
Public Const OBJ_BITMAP = 7      ' Returns the current selected bitmap
Public Const OBJ_BRUSH = 2       ' Returns the current selected brush
Public Const OBJ_COLORSPACE = 14 ' Returns the current color space
Public Const OBJ_FONT = 6        ' Returns the current selected font
Public Const OBJ_PAL = 5         ' Returns the current selected pal
Public Const OBJ_PEN = 1         ' Returns the current selected pen
' Constants - CopyImage.fuFlags
Public Const LR_COPYDELETEORG = &H8       ' Deletes the original image after creating the copy.
Public Const LR_COPYFROMRESOURCE = &H4000 ' Tries to reload an icon or cursor resource from the original resource file rather than simply copying the current image. This is useful for creating a different-sized copy when the resource file contains multiple sizes of the resource. Without this flag, CopyImage stretches the original image to the new size. If this flag is set, CopyImage uses the size in the resource file closest to the desired size.  This will succeed only if hImage was loaded by LoadIcon or LoadCursor, or by LoadImage with the LR_SHARED flag.
Public Const LR_COPYRETURNORG = &H4       ' Returns the original hImage if it satisfies the criteria for the copy—that is, correct dimensions and color depth—in which case the LR_COPYDELETEORG flag is ignored. If this flag is not specified, a new object is always created.
Public Const LR_CREATEDIBSECTION = &H2000 ' If this is set and a new bitmap is created, the bitmap is created as a DIB section. Otherwise, the bitmap image is created as a device-dependent bitmap. This flag is only valid if uType is IMAGE_BITMAP.
Public Const LR_MONOCHROME = &H1          ' Creates a new monochrome image.
' Constants - RedrawWindow.fuRedraw
Public Const RDW_ERASE = &H4
Public Const RDW_FRAME = &H400
Public Const RDW_INTERNALPAINT = &H2
Public Const RDW_INVALIDATE = &H1
Public Const RDW_NOERASE = &H20
Public Const RDW_NOFRAME = &H800
Public Const RDW_NOINTERNALPAINT = &H10
Public Const RDW_VALIDATE = &H8
Public Const RDW_ERASENOW = &H200
Public Const RDW_UPDATENOW = &H100
Public Const RDW_ALLCHILDREN = &H80
Public Const RDW_NOCHILDREN = &H40
' Constants - DrawIconEx.diFlags
Public Const DI_MASK = &H1        ' Performs the raster operation specified by ropMask.
Public Const DI_IMAGE = &H2       ' Performs the raster operation specified by ropImage.
Public Const DI_NORMAL = &H3      ' Combination of DI_IMAGE and DI_MASK.
Public Const DI_COMPAT = &H4      ' Draws the icon or cursor using the system default image rather than the user-specified image.
Public Const DI_DEFAULTSIZE = &H8 ' Draws the icon or cursor using the width and height specified by the system metric values for cursors or icons, if the cxWidth and cyWidth parameters are set to zero. If this flag is not specified and cxWidth and cyWidth are set to zero, the function uses the actual resource size.
' Constants - SetStretchBltMode.iStretchMode
Public Const BLACKONWHITE = 1                   ' Performs a Boolean AND operation using the color values for the eliminated and existing pixels. If the bitmap is a monochrome bitmap, this mode preserves black pixels at the expense of white pixels.
Public Const WHITEONBLACK = 2                   ' Performs a Boolean OR operation using the color values for the eliminated and existing pixels. If the bitmap is a monochrome bitmap, this mode preserves white pixels at the expense of black pixels.
Public Const COLORONCOLOR = 3                   ' Deletes the pixels. This mode deletes all eliminated lines of pixels without trying to preserve their information.
Public Const HALFTONE = 4                       ' Maps pixels from the source rectangle into blocks of pixels in the destination rectangle. The average color over the destination block of pixels approximates the color of the source pixels. After setting the HALFTONE stretching mode, an application must call the SetBrushOrgEx function to set the brush origin. If it fails to do so, brush misalignment occurs.
Public Const MAXSTRETCHBLTMODE = 4              ' (undocumented)
Public Const STRETCH_ANDSCANS = BLACKONWHITE    ' Same as BLACKONWHITE.
Public Const STRETCH_ORSCANS = WHITEONBLACK     ' Same as WHITEONBLACK.
Public Const STRETCH_DELETESCANS = COLORONCOLOR ' Same as COLORONCOLOR.
Public Const STRETCH_HALFTONE = HALFTONE        ' Same as HALFTONE.
' Win32 Function Declarations
Public Declare Function BitBlt Lib "GDI32.DLL" (ByVal hDC_Destination As Long, ByVal X_Dest As Long, ByVal Y_Dest As Long, ByVal Width_Dest As Long, ByVal Height_Dest As Long, ByVal hdc_source As Long, ByVal X_Src As Long, ByVal Y_Src As Long, ByVal RasterOperation As Long) As Long
Public Declare Function CopyCursor Lib "USER32.DLL" (ByVal pCursor As Long) As Long
Public Declare Function CopyImage Lib "USER32.DLL" (ByVal hImage As Long, ByVal uType As Long, ByVal OutputWidth As Long, ByVal OutputHeight As Long, ByVal fuFlags As Long) As Long
Public Declare Function CopyIcon Lib "USER32.DLL" (ByVal hIcon As Long) As Long
Public Declare Function CreateBitmap Lib "GDI32.DLL" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal cPlanes As Long, ByVal cBitsPerPel As Long, ByRef lpvBits As Any) As Long
Public Declare Function CreateCompatibleBitmap Lib "GDI32.DLL" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "GDI32.DLL" (ByVal hdc As Long) As Long
Public Declare Function CreateIconIndirect Lib "USER32.DLL" (ByRef pICONINFO As ICONINFO) As Long
Public Declare Function DeleteDC Lib "GDI32.DLL" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "GDI32.DLL" (ByVal hGDIObj As Long) As Long
Public Declare Function DestroyIcon Lib "USER32.DLL" (ByVal hIcon As Long) As Long
Public Declare Function DrawIconEx Lib "USER32.DLL" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long, ByVal IconWidth As Long, ByVal IconHeight As Long, ByVal AniFrameIndex As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function GetCurrentObject Lib "GDI32.DLL" (ByVal hdc As Long, ByVal uObjectType As Long) As Long
Public Declare Function GetDC Lib "USER32.DLL" (ByVal hWnd As Long) As Long
Public Declare Function GetDesktopWindow Lib "USER32.DLL" () As Long
Public Declare Function GetDeviceCaps Lib "GDI32.DLL" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function GetEnhMetaFileHeader Lib "GDI32.DLL" (ByVal hEnhancedMetafile As Long, ByVal BufferSize As Long, ByRef lpEMH As ENHMETAHEADER) As Long
Public Declare Function GetIconInfo Lib "USER32.DLL" (ByVal hIcon As Long, ByRef pICONINFO As ICONINFO) As Long
Public Declare Function GetMapMode Lib "GDI32.DLL" (ByVal hdc As Long) As Long
Public Declare Function GetObjectAPI Lib "GDI32.DLL" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetPixel Lib "GDI32.DLL" (ByVal hdc As Long, ByVal XPos As Long, ByVal nYPos As Long) As Long
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (ByRef PicDesc As Any, ByRef RefIID As GUID, ByVal fPictureOwnsHandle As Long, ByRef IPic As StdPicture) As Long 'As IPicture) As Long
Public Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pColorRef As Long) As Long
Public Declare Function RedrawWindow Lib "USER32.DLL" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function ReleaseDC Lib "USER32.DLL" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "GDI32.DLL" (ByVal hdc As Long, ByVal hGDIObj As Long) As Long
Public Declare Function SetBkColor Lib "GDI32.DLL" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetMapMode Lib "GDI32.DLL" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Public Declare Function SetPixel Lib "GDI32.DLL" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function SetStretchBltMode Lib "GDI32.DLL" (ByVal hdc As Long, ByVal iStretchMode As Long) As Long
Public Declare Function SetBrushOrgEx Lib "GDI32.DLL" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, ByRef lpPoint As Any) As Long
Public Declare Function StretchBlt Lib "GDI32.DLL" (ByVal hDC_Destination As Long, ByVal X_Dest As Long, ByVal Y_Dest As Long, ByVal New_Width As Long, ByVal New_Height As Long, ByVal hdc_source As Long, ByVal X_Src As Long, ByVal Y_Src As Long, ByVal Orig_Width As Long, ByVal Orig_Height As Long, ByVal RasterOperation As Long) As Long
' Convert_HM_PX
'
' When dealing with the "Picture" property of VB objects such as PictureBox or Form, or StdPicture
' objects, the Height & Width properties of such is not measured in Pixels or Twips... but in something
' called "HiMetric".  This function takes the height and width measurements of a picture in HiMetric
' and converts it to Pixels so that it can be used with standard Win32 API calls, or with VB objects
' that have their "ScaleMode" property set to "vbPixels"
'
' NOTE - You can also use the "GetBitmapInfo" function to get the height and/or width of a picture
' in pixels.
'
' Parameter:              Use:
' --------------------------------------------------
' InputHeight             Optional. Specifies the height of the picture in HiMetric
' InputWidth              Optional. Specifies the width of the picture in HiMetric
' OutputHeight            Optional. Returns the height of the picture in Pixels (if InputHeight is valid)
' OutputWidth             Optional. Returns the width of the picture in Pixels (if InputWidth is valid)
' VB_Picture              Optional. If set to TRUE, the calculation used to get the desired return value
'                         uses the Screen.TwipsPerPixel properties to get the TwipsPerPixel instead of
'                         using a more accurate calculation of TwipsPerPixel.  The return value is correct
'                         for use with the Picture property of VB objects like PictureBox, & StdPicture.
'                         If set to FALSE, the calculation used to get the desired return value uses a
'                         calculation to get and use a more accurate measurement of the TwipsPerPixel.
'                         This is more accurate for use with Win32 API calls.
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
'
' --------------------------------------------------
' These type definitions were taken from OCIDL.H
' --------------------------------------------------
' typedef LONG OLE_XPOS_HIMETRIC;
' typedef LONG OLE_YPOS_HIMETRIC;
' typedef LONG OLE_XSIZE_HIMETRIC;
' typedef LONG OLE_YSIZE_HIMETRIC;
'
Public Function Convert_HM_PX(Optional ByVal InputHeight As Long, Optional ByVal InputWidth As Long, Optional ByRef OutputHeight As Long, Optional ByRef OutputWidth As Long, Optional ByVal VB_Picture As Boolean = True) As Boolean
   On Error Resume Next
   Dim TwipsX As Single
   Dim TwipsY As Single
   ' Reset the return values
   OutputHeight = 0
   OutputWidth = 0
   ' Make sure the parameters passed are valid
   If InputHeight = 0 And InputWidth = 0 Then Exit Function
   ' If the user specifies to do the convertion for a Visual Basic Picture, use the
   ' "Screen" object to get the approximate TwipsPerPixel
   If VB_Picture = True Then
      OutputWidth = CLng(((InputWidth / 2540) * 1440) / Screen.TwipsPerPixelX)
      OutputHeight = CLng(((InputHeight / 2540) * 1440) / Screen.TwipsPerPixelY)
      ' If the user doesn't specify to do the convertion for a Visual Basic Picture, assume
      ' it's for a Win32 API call and calculate the exact TwipsPerPixel to be more accurate
   Else
      If GetDisplayInfo(, , TwipsX, TwipsY) = False Then Exit Function
      OutputWidth = CLng((InputWidth / 2540 * 1440) / TwipsX)
      OutputHeight = CLng((InputHeight / 2540 * 1440) / TwipsY)
   End If
   ' Function succeeded
   Convert_HM_PX = True
End Function
' Convert_PX_HM
'
' When dealing with the "Picture" property of VB objects such as PictureBox or Form, or StdPicture
' objects, the Height & Width properties of such is not measured in Pixels or Twips... but in something
' called "HiMetric".  This function takes the height and width measurements of a picture in Pixels
' and converts it to HiMetric so that it can be used with VB calls, etc.
'
' NOTE - When the "VB_Picture" parameter is set to FALSE, the return values of this function are VERY
' close, but not exact because of how the calculations and number rounding works.  To see this effect,
' use the Convert_HM_PX function to take the height/width of a picture and convert them to pixels...
' then take the return values from that and use this function to convert them back to their original
' HiMetrics measurement.  The results will be very close, but not exact.  This shouldn't be a problem
' because I would think it would be a rare thing that you'd want to convert Pixels to HiMetric (I even
' considered leaving this function out of the module).
'
' Parameter:              Use:
' --------------------------------------------------
' InputHeight             Optional. Specifies the height of the picture in Pixels
' InputWidth              Optional. Specifies the width of the picture in Pixels
' OutputHeight            Optional. Returns the height of the picture in HiMetric (if InputHeight is valid)
' OutputWidth             Optional. Returns the width of the picture in HiMetric (if InputWidth is valid)
' VB_Picture              Optional. If set to TRUE, the calculation used to get the desired return value
'                         uses the Screen.TwipsPerPixel properties to get the TwipsPerPixel instead of
'                         using a more accurate calculation of TwipsPerPixel.  The return value is correct
'                         for use with the Picture property of VB objects like PictureBox, & StdPicture.
'                         If set to FALSE, the calculation used to get the desired return value uses a
'                         calculation to get and use a more accurate measurement of the TwipsPerPixel.
'                         This is more accurate for use with Win32 API calls.
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
'
' --------------------------------------------------
' These type definitions were taken from OCIDL.H
' --------------------------------------------------
' typedef LONG OLE_XPOS_HIMETRIC;
' typedef LONG OLE_YPOS_HIMETRIC;
' typedef LONG OLE_XSIZE_HIMETRIC;
' typedef LONG OLE_YSIZE_HIMETRIC;
'
Public Function Convert_PX_HM(ByVal InputHeight As Long, ByVal InputWidth As Long, ByRef OutputHeight As Long, ByRef OutputWidth As Long, Optional ByVal VB_Picture As Boolean = True) As Boolean
   On Error Resume Next
   Dim TwipsX As Single
   Dim TwipsY As Single
   ' Reset the return values
   OutputHeight = 0
   OutputWidth = 0
   ' Make sure the parameters passed are valid
   If InputHeight = 0 And InputWidth = 0 Then Exit Function
   ' If the user specifies to do the convertion for a Visual Basic Picture, use the
   ' "Screen" object to get the approximate TwipsPerPixel
   If VB_Picture = True Then
      OutputHeight = CLng(((InputHeight * Screen.TwipsPerPixelY) / 1440) * 2540)
      OutputWidth = CLng(((InputWidth * Screen.TwipsPerPixelX) / 1440) * 2540)
      ' If the user doesn't specify to do the convertion for a Visual Basic Picture, assume
      ' it's for a Win32 API call and calculate the exact TwipsPerPixel to be more accurate
   Else
      If GetDisplayInfo(, , TwipsX, TwipsY) = False Then Exit Function
      OutputHeight = CLng(((InputHeight * TwipsX) / 1440) * 2540)
      OutputWidth = CLng(((InputWidth * TwipsY) / 1440) * 2540)
   End If
   ' Function succeeded
   Convert_PX_HM = True
End Function
'
' CopyPicture
'
' This function takes the handle to the picture passed in via the "IN_hPicture" parameter and makes a
' copy of it... returning it via the "OUT_hPicture" parameter.
'
' Parameter:              Use:
' --------------------------------------------------
' IN_hPicture             Specifies the handle to the picture to copy
' OUT_hPicture            Returns the newly created copy of the original picture
' PictureType             Optional. Specifies the type of image to copy (Bitmap, Icon, Cursor, Enh Metafile)
' PictureWidth            Optional. Specifies the width of the image to copy.  If this is not specified,
'                         this function attempts to get the width from the image.
' PictureHeight           Optional. Specifies the height of the image to copy.  If this is not specified,
'                         this function attempts to get the height from the image.
' ReturnMonochrome        Optional. If set to TRUE, the return is a black and white version of the image
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
'
Public Function CopyPicture(ByVal IN_hPicture As Long, ByRef OUT_hPicture As Long, Optional ByVal PictureType As PictureTypes = IMAGE_BITMAP, Optional ByVal PictureWidth As Long, Optional ByVal PictureHeight As Long, Optional ByVal ReturnMonochrome As Boolean = False) As Boolean
   Dim TempEMH     As ENHMETAHEADER
   Dim TempBITMAP  As BITMAP
   Dim hBMP_Mask   As Long
   Dim hBMP_Image  As Long
   Dim ReturnValue As Long
   Dim flags       As Long
   ' Set the default return value
   OUT_hPicture = 0
   ' Make sure parameters passed are valid
   If IN_hPicture = 0 Then Exit Function
   ' Get the dimentions and type of picture to copy
   If PictureWidth = 0 Or PictureHeight = 0 Then
      Select Case PictureType
      Case IMAGE_BITMAP
         If GetObjectAPI(IN_hPicture, Len(TempBITMAP), TempBITMAP) = 0 Then Exit Function
         PictureWidth = TempBITMAP.bmWidth
         PictureHeight = TempBITMAP.bmHeight
      Case IMAGE_ICON, IMAGE_CURSOR
         If GetIconBitmaps(IN_hPicture, hBMP_Mask, hBMP_Image) = False Then Exit Function
         ReturnValue = GetObjectAPI(hBMP_Image, Len(TempBITMAP), TempBITMAP)
         DeleteObject hBMP_Mask
         DeleteObject hBMP_Image
         If ReturnValue = 0 Then Exit Function
         PictureWidth = TempBITMAP.bmWidth
         PictureHeight = TempBITMAP.bmHeight
      Case IMAGE_ENHMETAFILE
         TempEMH.nSize = Len(TempEMH)
         TempEMH.iType = EMR_HEADER
         TempEMH.dSignature = ENHMETA_SIGNATURE
         TempEMH.nVersion = &H10000
         If GetEnhMetaFileHeader(IN_hPicture, Len(TempEMH), TempEMH) = 0 Then Exit Function
         PictureWidth = TempEMH.rclBounds.right
         PictureHeight = TempEMH.rclBounds.bottom
      End Select
   End If
   ' Copy the image
   If ReturnMonochrome = True Then flags = LR_MONOCHROME
   OUT_hPicture = CopyImage(IN_hPicture, CLng(PictureType), PictureWidth, PictureHeight, flags)
   If OUT_hPicture <> 0 Then CopyPicture = True
End Function
' CreateCursorFromBMP
'
' This function takes the handle to the mask and image BITMAPS that make up an cursor, and combine them
' to make a transparent icon.
'
' Parameter:              Use:
' --------------------------------------------------
' hBMP_Mask               Handle to the mask BITMAP to use
' hBMP_Image              Handle to the image BITMAP to use
'
' Return:
' -------
' If the function succeeds, the return is the handle to the newly created icon
' If the function fails, the return is ZERO (0)
'
Public Function CreateCursorFromBMP(ByVal hBMP_Mask As Long, ByVal hBMP_Image As Long, Optional ByVal HotspotX As Long, Optional ByVal HotspotY As Long) As Long
   Dim TempICONINFO As ICONINFO
   If hBMP_Mask = 0 Or hBMP_Image = 0 Then Exit Function
   TempICONINFO.fIcon = 0
   TempICONINFO.hbmMask = hBMP_Mask
   TempICONINFO.hbmColor = hBMP_Image
   TempICONINFO.xHotspot = HotspotX
   TempICONINFO.yHotspot = HotspotY
   CreateCursorFromBMP = CreateIconIndirect(TempICONINFO)
End Function
'
' CreateIconFromBMP
'
' This function takes the handle to the mask and image BITMAPS that make up an icon, and combine them
' to make a transparent icon.
'
' Parameter:              Use:
' --------------------------------------------------
' hBMP_Mask               Handle to the mask BITMAP to use
' hBMP_Image              Handle to the image BITMAP to use
'
' Return:
' -------
' If the function succeeds, the return is the handle to the newly created icon
' If the function fails, the return is ZERO (0)
'
Public Function CreateIconFromBMP(ByVal hBMP_Mask As Long, ByVal hBMP_Image As Long) As Long
   Dim TempICONINFO As ICONINFO
   If hBMP_Mask = 0 Or hBMP_Image = 0 Then Exit Function
   TempICONINFO.fIcon = 1
   TempICONINFO.hbmMask = hBMP_Mask
   TempICONINFO.hbmColor = hBMP_Image
   CreateIconFromBMP = CreateIconIndirect(TempICONINFO)
End Function
' CreateMask
'
' This function takes the specified picture and creates a sprite and a mask from it.  The sprite is the
' same as the original picture, but the color that is specified by the "TransparentColor" parameter is
' changed to WHITE (this serves to designate where the transparency will be).  The mask is a black
' silhouette of the original picture with a white background.
'
' When the mask is combined with another picture using the Win32 "BitBlt" API with the "MERGEPAINT"
' raster operation, it puts a white silhouette of the original picture (without the transparent region).
' When the sprite is combined with the picture that the mask was combined with in the same location
' as the mask using the Win32 "BitBlt" API with the "SRCAND" raster operation, the original picture is
' displayed on the picture as a transparent picture (the specified background color, or transparent
' color no longer shows up.
'
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' Dim PicBG      As StdPicture
' Dim PicImg     As StdPicture
' Dim hSpriteDC  As Long
' Dim hMaskDC    As Long
' Dim hSpriteBMP As Long
' Dim hMaskBMP   As Long
' Dim TheWidth   As Long
' Dim TheHeight  As Long
'
' Me.Show
' Set PicBG = LoadPicture("C:\Background.bmp")
' Set PicImg = LoadPicture("C:\Image.bmp")
' RenderBitmapEx  Me.hDC, , PicBG.Handle
' GetBitmapInfo PicImg.Handle, TheHeight, TheWidth
' CreateMask PicImg.Handle, CLng("&H00FF00"), hSpriteDC, hMaskDC, hSpriteBMP, hMaskBMP
' BitBlt Me.hDC, 0, 0, TheWidth, TheHeight, hMaskDC, 0, 0, MERGEPAINT
' BitBlt Me.hDC, 0, 0, TheWidth, TheHeight, hSpriteDC, 0, 0, SRCAND
' MemoryDC_Delete hSpriteDC, hSpriteBMP
' MemoryDC_Delete hMaskDC, hMaskBMP
' Set PicBG = Nothing
' Set PicImg = Nothing
'
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'
' Parameter:              Use:
' --------------------------------------------------
' hBitmap                 Handle to the bitmap to create the sprite and mask from
' TransparentColor        Specifies the color that is to be made transparent (background color)
' Return_SpriteDC         Optional. Returns the handle to the DC created that contains the sprite
' Return_MaskDC           Optional. Returns the handle to the DC created that contains the mask
' Return_SpritePrevBMP    Optional. Returns the handle to the picture that was previously in the sprite DC.
'                         This is important because you should always select the OLD picture back into
'                         a DC before deleting it.
' Return_MaskPrevBMP      Optional. Returns the handle to the picture that was previously in the mask DC.
'                         This is important because you should always select the old picture back into
'                         a DC before deleting it.
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
'
Public Function CreateMask(ByVal hBitmap As Long, ByVal TransparentColor As Long, Optional ByRef Return_SpriteDC As Long = -1, Optional ByRef Return_MaskDC As Long = -1, Optional ByRef Return_SpritePrevBMP As Long = -1, Optional ByRef Return_MaskPrevBMP As Long = -1) As Boolean
   Dim hScreenDC     As Long
   Dim PicHeight     As Long
   Dim PicWidth      As Long
   Dim hDC_Temp      As Long
   Dim hDC_Sprite    As Long
   Dim hDC_Mask      As Long
   Dim hPrev_Temp    As Long
   Dim hPrev_Sprite  As Long
   Dim hPrev_Mask    As Long
   Dim PreviousColor As Long
   Dim hMonoBitmap   As Long
   Dim hColrBitmap   As Long
   ' Make sure a valid picture was passed to use to get the mask
   If hBitmap = 0 Then Exit Function
   ' If the user hasn't specified to return any values, exit
   If Return_SpriteDC = -1 And Return_MaskDC = -1 Then Exit Function
   ' Get the transparent color as a non-OLE color
   TransparentColor = TranslateColor(TransparentColor)
   If TransparentColor = -1 Then Exit Function
   ' Get the height & width of the picture
   If GetBitmapInfo(hBitmap, PicHeight, PicWidth) = False Then Exit Function
   ' Create the Device Contexts (DC) to work with
   hDC_Temp = MemoryDC_Create
   hDC_Mask = MemoryDC_Create
   hDC_Sprite = MemoryDC_Create
   ' Create a monochrome bitmap to create the mask from
   hMonoBitmap = CreateCompatibleBitmap(hDC_Temp, PicWidth, PicHeight)
   ' Create a color bitmap to select into the mask
   hScreenDC = GetDC(GetDesktopWindow)
   hColrBitmap = CreateCompatibleBitmap(hScreenDC, PicWidth, PicHeight)
   ReleaseDC GetDesktopWindow, hScreenDC
   ' Copy the sprite picture into the sprite DC making it color, then
   ' copy the monochrome bitmaps into the temp and mask DC's black & white
   LoadBitmapToDC hDC_Temp, hMonoBitmap, hPrev_Temp
   LoadBitmapToDC hDC_Mask, hColrBitmap, hPrev_Mask
   LoadBitmapToDC hDC_Sprite, hBitmap, hPrev_Sprite
   ' Set the background color to the transparent color.
   ' This will make the transparent color black when the mask is created
   PreviousColor = SetBkColor(hDC_Sprite, TransparentColor)
   ' Copy the sprite picture to the temp DC - This makes a copy of the
   ' sprite, but the copy is black and the transparent color is now transparent
   BitBlt hDC_Temp, 0, 0, PicWidth, PicHeight, hDC_Sprite, 0, 0, SRCCOPY
   ' Set the background color back to the original color
   SetBkColor hDC_Sprite, PreviousColor
   ' Copy the temp DC to the sprite - this makes the sprite's transparent
   ' color now show up as WHITE
   BitBlt hDC_Sprite, 0, 0, PicWidth, PicHeight, hDC_Temp, 0, 0, SRCPAINT
   ' Create the mask (Silhouette of the original on a white background)
   BitBlt hDC_Mask, 0, 0, PicWidth, PicHeight, hDC_Sprite, 0, 0, SRCCOPY
   BitBlt hDC_Mask, 0, 0, PicWidth, PicHeight, hDC_Temp, 0, 0, SRCCOPY
   ' Return the results and clean up extra memory
   CreateMask = True
   ' If the user specified to return the sprite, do so... else delete it
   If Return_SpriteDC = -1 Then
      MemoryDC_Delete hDC_Sprite, hPrev_Sprite
      Return_SpriteDC = 0
      Return_SpritePrevBMP = 0
   Else
      Return_SpriteDC = hDC_Sprite
      If Return_SpritePrevBMP = -1 Then
         DeleteObject hPrev_Sprite
      Else
         Return_SpritePrevBMP = hPrev_Sprite
      End If
   End If
   ' If the user specified to return the mask, do so... else delete it
   If Return_MaskDC = -1 Then
      MemoryDC_Delete hDC_Mask, hPrev_Mask
      Return_MaskDC = 0
      Return_MaskPrevBMP = 0
   Else
      Return_MaskDC = hDC_Mask
      If Return_MaskPrevBMP = -1 Then
         DeleteObject hPrev_Mask
      Else
         Return_MaskPrevBMP = hPrev_Mask
      End If
   End If
   ' Clean up memory used to create the sprite and mask
   DeleteObject hMonoBitmap: hMonoBitmap = 0
   DeleteObject hColrBitmap: hColrBitmap = 0
   If hDC_Temp <> 0 Then MemoryDC_Delete hDC_Temp, hPrev_Temp
End Function
' CreateOlePicture
'
' This function takes the handle to a picture (Bitmap, Icon, Metafile, or Enhanced Metafile) and creates
' an OLE StdPicture object from it that can be used like the "Picture" properties of such VB objects as
' Form's, PictureBox's, ImageBox's, etc.
'
' Parameter:              Use:
' --------------------------------------------------
' PictureHandle           Handle to the picture to create.
'                           - If PictureType = vbPicTypeBitmap    : this must be a handle to a HBITMAP
'                           - If PictureType = vbPicTypeIcon      : this must be a handle to a HICON
'                           - If PictureType = vbPicTypeMetafile  : this must be a handle to a HMETAFILE
'                           - If PictureType = vbPicTypeEMetafile : this must be a handle to a HENHMETAFILE
' PictureType             Specifies the type of picture object to create.  These are the different types
'                         of pictures that can be specified:
'                           vbPicTypeBitmap     <-- DEFAULT
'                           vbPicTypeEMetafile
'                           vbPicTypeIcon
'                           vbPicTypeMetafile
' BitmapPalette           Optional. Specifies the handle to a Palette to use in the createion process.
' MetaHeight              Optional. If the PictureType is vbPicTypeMetafile, the height of the Metafile
'                         must be provided by this parameter.
' Metawidth               Optional. If the PictureType is vbPicTypeMetafile, the width of the Metafile
'                         must be provided by this parameter.
' Return_ErrNum           Optional. If an error occurs, the error number will be returned here.
' Return_ErrDesc          Optional. If an error occurs, the error description will be returned here.
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
'
Public Function CreateOlePicture(ByVal PictureHandle As Long, ByVal PictureType As PictureTypeConstants, Optional ByVal BitmapPalette As Long = 0, Optional ByVal MetaHeight As Long = -1, Optional ByVal MetaWidth As Long = -1, Optional ByRef Return_ErrNum As Long, Optional ByRef Return_ErrDesc As String) As StdPicture
   On Error Resume Next
   Dim ReturnValue   As Long
   Dim PicInfo_BMP   As PICTDESC_BMP
   Dim PicInfo_EMETA As PICTDESC_EMETA
   Dim PicInfo_ICON  As PICTDESC_ICON
   Dim PicInfo_META  As PICTDESC_META
   Dim ThePicture    As StdPicture 'IPicture
   Dim rIID          As GUID
   ' Clear the return variables
   Return_ErrNum = 0
   Return_ErrDesc = ""
   ' Make sure the variable(s) passed are valid
   If PictureHandle = 0 Then
      Return_ErrNum = -1
      Return_ErrDesc = "Invalid bitmap handle"
   ElseIf PictureType = vbPicTypeNone Then
      Return_ErrNum = -1
      Return_ErrDesc = "Invalid picture type specified."
   ElseIf PictureType = vbPicTypeMetafile Then
      If MetaHeight = -1 Or MetaWidth = -1 Then
         Return_ErrNum = -1
         Return_ErrDesc = "Invalid metafile dimentions specified."
      End If
   End If
   ' Set the correct interface identifier GUID for the "OleCreatePictureIndirect" API
   With rIID
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With
   ' Set the appropriate type depending on the type of picture
   Select Case PictureType
   Case vbPicTypeBitmap
      PicInfo_BMP.cbSizeOfStruct = Len(PicInfo_BMP)
      PicInfo_BMP.PicType = PICTYPE_BITMAP
      PicInfo_BMP.hBitmap = PictureHandle
      PicInfo_BMP.hPal = BitmapPalette
      ReturnValue = OleCreatePictureIndirect(PicInfo_BMP, rIID, 1, ThePicture)
   Case vbPicTypeIcon
      PicInfo_ICON.cbSizeOfStruct = Len(PicInfo_BMP)
      PicInfo_ICON.PicType = PICTYPE_ICON
      PicInfo_ICON.hIcon = PictureHandle
      ReturnValue = OleCreatePictureIndirect(PicInfo_ICON, rIID, 1, ThePicture)
   Case vbPicTypeMetafile
      PicInfo_META.cbSizeOfStruct = Len(PicInfo_BMP)
      PicInfo_META.PicType = PICTYPE_METAFILE
      PicInfo_META.hMeta = PictureHandle
      PicInfo_META.xExt = MetaWidth
      PicInfo_META.yExt = MetaHeight
      ReturnValue = OleCreatePictureIndirect(PicInfo_META, rIID, 1, ThePicture)
   Case vbPicTypeEMetafile
      PicInfo_EMETA.cbSizeOfStruct = Len(PicInfo_BMP)
      PicInfo_EMETA.PicType = PICTYPE_ENHMETAFILE
      PicInfo_EMETA.hEMF = PictureHandle
      ReturnValue = OleCreatePictureIndirect(PicInfo_BMP, rIID, 1, ThePicture)
   End Select
   ' Check the result
   If ReturnValue <> S_OK Then
      GoTo ErrorTrap
   End If
   ' Return the new picture
   Set CreateOlePicture = ThePicture
   Exit Function
ErrorTrap:
   Return_ErrNum = ReturnValue
   Select Case ReturnValue
   Case E_NOINTERFACE
      Return_ErrDesc = "The object does not support the interface specified in riid."
   Case E_POINTER
      Return_ErrDesc = "The address in pPictDesc or ppvObj is not valid. For example, it may be NULL."
   Case E_INVALIDARG
      Return_ErrDesc = "One or more arguments are invalid."
   Case E_OUTOFMEMORY
      Return_ErrDesc = "Ran out of memory."
   Case E_UNEXPECTED
      Return_ErrDesc = "Catastrophic Failure."
   Case Else
      Return_ErrDesc = "Unknown Error."
   End Select
End Function
'
' ExtractFromRES
'
' This function extracts the specified files from a VB resource that is included in your project.
'
' Parameter:              Use:
' --------------------------------------------------
' OutputFile              Specifies the name of the file to extract the resource to.  If the file
'                         specified already exists, and the "OverwriteFile" parameter is set to TRUE,
'                         the file will be automatically overwritten.  If the "OverwriteFile" parameter
'                         is FALSE and the "ConfirmOnOverwrite" parameter is TRUE, the user is prompted
'                         to overwrite.  If the file exists, the OverwriteFile parameter is FALSE, and
'                         the ConfirmOnOverwrite parameter is FALSE, the function fails out.
'                         NOTE: If you specify "vbNullString" or a blank string ("") for this parameter
'                         and specify TRUE for the "ReturnPicRef" parameter, the file will not be saved
'                         out, but a reference to it will be returned via the "Return_Picture" parameter.
' RES_ID                  Specifies the resource ID of the file to extract.  The default ID is 101.
' RES_Section             Specifies the section of resource to extract from.  When you add a file to a
'                         VB resource, the default type is "CUSTOM".  You can of course change the
'                         type of the file added in the resource editor.  This parameter is only used
'                         if the "FileType" parameter is rt_Custom.
' FileType                Specifies if the file to extract is a Bitmap, Icon, Cursor, or other file.
' OverwriteFile           If the specified file already exists and this parameter is TRUE, the file is
'                         overwritten.
' ConfirmOnOverwrite      If the file exists, and OverwriteFile is FALSE, and this parameter is TRUE,
'                         the user is prompted to overwrite the file.
' ReturnPicRef            If set to TRUE, the "Return_Picture" parameter returns a reference to the
'                         picture specified.  If the FileType is set to rt_Custom, and this parameter
'                         is set to true, it is assumed that there is a picture stored int he custom
'                         resource and this function attempts to load it.
' Return_Picture          If the "ReturnPicRef" parameter is set to TRUE, this parameter returns a
'                         reference to the picture loaded.  If the FileType is set to rt_Custom, and
'                         the "ReturnPicRef" parameter is set to true, it is assumed that there is a
'                         picture stored int he custom resource and this function attempts to load it.
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
'
Public Function ExtractFromRES(ByVal OutputFile As String, ByVal RES_ID As Long, Optional ByVal RES_Section As String, Optional ByVal FileType As ResTypes = rt_Custom, Optional ByVal OverwriteFile As Boolean = True, Optional ByVal ConfirmOnOverwrite As Boolean = False, Optional ByVal ReturnPicRef As Boolean = False, Optional ByRef Return_Picture As StdPicture) As Boolean
   On Error GoTo ErrorTrap
   Dim ResFile() As Byte
   Dim TestVar   As Variant
   Dim RESPic    As StdPicture
   Dim FileNum   As Long
   Dim MyAnswer  As VbMsgBoxResult
   Dim DelFile   As Boolean
   ' Make sure parameters are valid
   If RES_ID < 1 Or RES_ID > 32767 Then
      MsgBox "Resource ID is invalid.  Value must be between 1 and 32767.", vbOKOnly + vbExclamation, "  Invalid Resource ID"
      Exit Function
   ElseIf FileType = rt_Custom And RES_Section = "" Then
      MsgBox "No resource type specified.", vbOKOnly + vbExclamation, "  Error Extracting Resource"
      Exit Function
   ElseIf OutputFile = "" And ReturnPicRef = False Then
      MsgBox "No output file specified to extract resource to.", vbOKOnly + vbExclamation, "  No Output File Specified"
      Exit Function
   ElseIf Dir(OutputFile) <> "" And OverwriteFile = False Then
      If ConfirmOnOverwrite = True Then
         MyAnswer = MsgBox(OutputFile & VBA.Chr(13) & "This file already exists." & VBA.Chr(13) & VBA.Chr(13) & "Overwrite existing file?", vbYesNo + vbExclamation, "  Confirm File Overwrite")
         If MyAnswer = vbNo Then
            ExtractFromRES = True
            Exit Function
         End If
      Else
         Exit Function
      End If
   End If
   ' Extract and process the specified file
   Select Case FileType
      ' Save out the specified picture resource
   Case RT_BITMAP, RT_CURSOR, RT_ICON
      Set RESPic = LoadResPicture(RES_ID, FileType)
      If RESPic Is Nothing Then Exit Function
      If OutputFile <> "" Then SavePicture RESPic, OutputFile
      If ReturnPicRef = True Then
         Set Return_Picture = RESPic
      Else
         Set RESPic = Nothing
      End If
      ' Save out the specified custom resource file
   Case rt_Custom
      ' If no output file specified, use a temporary one
      If OutputFile = "" Then
         OutputFile = "C:\TEMP.TMP"
         On Error Resume Next
         Kill OutputFile
         Err.Clear
         On Error GoTo ErrorTrap
         DelFile = True
      End If
      ' Load the resource
      ResFile = LoadResData(RES_ID, RES_Section)
      ' Check if an error occured while loading the resource
      On Error Resume Next
      Err.Clear
      TestVar = UBound(ResFile)
      If Err Then Err.Clear: Exit Function
      On Error GoTo ErrorTrap
      ' Save the resource file out to disk
      FileNum = FreeFile
      Open OutputFile For Binary As #FileNum
      Put #FileNum, 1, ResFile()
      Close #FileNum
      ' If the user specified to return the picture, assume the custom file is a picture
      ' and try to load it and return a reference to it
      If ReturnPicRef = True Then
         On Error Resume Next
         Set Return_Picture = LoadPicture(OutputFile)
         If Err Or (Return_Picture Is Nothing) Then
            Set Return_Picture = Nothing
            Err.Clear
         End If
      End If
      ' Clean up temp file
      If DelFile = True Then Kill OutputFile
   End Select
   ' Return SUCCESS
   ExtractFromRES = True
   Exit Function
ErrorTrap:
   If Err.Number = 0 Then      ' No Error
      Resume Next
   ElseIf Err.Number = 20 Then ' Resume Without Error
      Resume Next
   Else                        ' Unknown Error
      MsgBox Err.Source & " encountered the following error:" & VBA.Chr(13) & VBA.Chr(13) & "Error Number = " & CStr(Err.Number) & VBA.Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
      Err.Clear
   End If
End Function
'
' GetDisplayInfo
'
' This function returns general information about the current screen display settings.
'
' Parameter:              Use:
' --------------------------------------------------
' ScreenWidth             Returns the width of the current display resolution in pixels
' ScreenHeight            Returns the height of the current display resolution in pixels
' TwipsX                  Returns the current number of twips per each pixel accross the X axis of the
'                         current screen display (see also Screen.TwipsPerPixelX)
' TwipsY                  Returns the current number of twips per each pixel accross the Y axis of the
'                         current screen display (see also Screen.TwipsPerPixelY)
' BitsPerPixel            Returns the current color depth of the screen display
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
'
Public Function GetDisplayInfo(Optional ByRef ScreenWidth As Long, Optional ByRef ScreenHeight As Long, Optional ByRef TwipsX As Single, Optional ByRef TwipsY As Single, Optional ByRef BitsPerPixel As Long) As Boolean
   On Error Resume Next
   Dim TempW     As Long
   Dim TempH     As Long
   Dim hScreenDC As Long
   ' Reset the return values
   ScreenWidth = 0
   ScreenHeight = 0
   TwipsX = 0
   TwipsY = 0
   BitsPerPixel = 0
   ' Get the handle to the display area's Device Context
   hScreenDC = GetDC(GetDesktopWindow)
   If hScreenDC = 0 Then Exit Function
   ' Get the color depth
   BitsPerPixel = GetDeviceCaps(hScreenDC, BITSPIXEL)
   ' Get the screen width / height in pixels
   ScreenWidth = GetDeviceCaps(hScreenDC, HORZRES)
   ScreenHeight = GetDeviceCaps(hScreenDC, VERTRES)
   ' Get the physical width / height of the screen in millimeters
   TempH = GetDeviceCaps(hScreenDC, VERTSIZE)
   TempW = GetDeviceCaps(hScreenDC, HORZSIZE)
   ' Get the TwipsPerPixelX & TwipsPerPixelY (There are 56.7 twips per millimeter)
   TwipsX = CSng((56.7 * TempW) / ScreenWidth)
   TwipsY = CSng((56.7 * TempH) / ScreenHeight)
   ' Format the return to one decimal place
   TwipsX = CSng(Format(TwipsX, "0.0"))
   TwipsY = CSng(Format(TwipsY, "0.0"))
   ' Clean up and return
   ReleaseDC GetDesktopWindow, hScreenDC
   GetDisplayInfo = True
End Function
' GetBitmapInfo
'
' This function takes a given picture and finds out all possible information about it and returns the
' results.
'
' NOTE : This function only works with BITMAPs and DIBs (Device Independant Bitmaps)
'
' Parameter:              Use:
' --------------------------------------------------
' hBITMAP                 Handle to the bitmap to get the information from.
' Return_Height           Optional. Returns the height (in pixels) of the picture.
' Return_Width            Optional. Returns the width (in pixels) of the picture.
' Return_BitsPerPixel     Optional. Returns the color depth of the picture in the form of "BitsPerPixel"
' Return_Size             Optional. Returns the approximate size of the picture (assuming it's RGB)
' Return_PointerToBits    Optional. Returns a memory pointer to the location of the BITMAP BITS that make
'                         up the specified image.  You can use the "CopyMemory" API to copy the BITS to a
'                         BYTE ARRAY.
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
'
Public Function GetBitmapInfo(ByVal hBitmap As Long, Optional ByRef Return_Height As Long, Optional ByRef Return_Width As Long, Optional ByRef Return_BitsPerPixel As Integer, Optional ByRef Return_Size As Double, Optional ByRef Return_PointerToBits As Long) As Boolean
   On Error Resume Next
   Dim BMP As BITMAP
   ' Clear the return variables
   Return_Height = 0
   Return_Width = 0
   Return_BitsPerPixel = 0
   Return_Size = 0
   Return_PointerToBits = 0
   ' Check that there's a valid input
   If hBitmap = 0 Then Exit Function
   ' Get the information
   If GetObjectAPI(hBitmap, Len(BMP), BMP) = 0 Then Exit Function
   ' Return the information
   With BMP
      Return_Height = .bmHeight
      Return_Width = .bmWidth
      Return_BitsPerPixel = (.bmBitsPixel * .bmPlanes)
      Return_Size = ((.bmWidth * 3 + 3) And &HFFFFFFFC) * .bmHeight
      Return_PointerToBits = .bmBits
   End With
   ' Function succeeded
   GetBitmapInfo = True
End Function
' GetIconBitmaps
'
' This function takes the given ICON or CURSOR and breaks out the Image and Mask BITMAPs that make it up.
' When you BitBlt the mask BITMAP onto a Device Context using the "SRCCOPY" raster operation, then BitBlt
' the image BITMAP onto the same Device Context in the same location as the mask using the "SRCINVERT"
' raster operation the result is a transparent picture.
'
' WARNING : The caller of this function is responsible for deleting the BITMAPs that this
'           function returns by calling the "DeleteObject" Win32 API.
'
' Parameter:              Use:
' --------------------------------------------------
' hIcon                   Handle to the icon to break the BITMAPs out of
' Return_hBmpMask         Returns the handle of the mask BITMAP (the caller must delete this BITMAP)
' Return_hBmpImage        Returns the handle of the image BITMAP (the caller must delete this BITMAP)
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
'
Public Function GetIconBitmaps(ByVal hIcon As Long, ByRef Return_hBmpMask As Long, ByRef Return_hBmpImage As Long) As Boolean
   Dim TempICONINFO As ICONINFO
   If GetIconInfo(hIcon, TempICONINFO) = 0 Then Exit Function
   Return_hBmpMask = TempICONINFO.hbmMask
   Return_hBmpImage = TempICONINFO.hbmColor
   GetIconBitmaps = True
End Function
' GetBitmapFromDC
'
' This function returns the handle to the picture that is currently selected into the specified Device
' Context (DC).  This does not remove the picture from the DC, it just gives you the handle to the picture
' so you can check the picture's height, width, color depth, etc. if you wish.
'
' Parameter:              Use:
' --------------------------------------------------
' hDC                     Handle to the DC to get the picture from
'
' Return:
' -------
' If the function succeeds, the return is the handle to the DC's BITMAP image
' If the function fails, the return is ZERO (0)
'
Public Function GetBitmapFromDC(ByVal hdc As Long) As Long
   GetBitmapFromDC = GetCurrentObject(hdc, OBJ_BITMAP)
End Function
' MemoryDC_Create
'
' This function creates a Device Context (DC) in memory compatible with the current screen display.
' This DC can be used to hold or render Bitmaps, Brushes, Fonts, Pens, or Regions (see also : Win32 API
' documentation for the "SelectObject" function).
'
' NOTE : Before an application can use a memory device context for drawing operations, it must select
' a bitmap of the correct width and height into the device context.  If the bitmap selected into the
' DC is MONOCHROME (created using the "CreateCompatibleBitmap" API with the handle to the newly created
' DC used as the hDC parameter for the "CreateCompatibleBitmap" API), then the DC is monochrome...
' otherwise it is color (compatible to the current screen).
'
' IMPORTANT: When you are done with the DC created by this function, you must delete it by calling the
' MemoryDC_Delete function in this module, or the Win32 API "DeleteDC".  If you do not, you could cause
' a VERY LARGE memory leak in your program.
'
' Parameter:              Use:
' --------------------------------------------------
' None
'
' Return:
' -------
' If the function succeeds, the return is the handle to the newly created Device Context
' If the function fails, the return is ZERO (0)
'
Public Function MemoryDC_Create() As Long
   Dim hScreenDC As Long
   ' Get the handle to the current display screen's DC
   hScreenDC = GetDC(GetDesktopWindow)
   ' Create a compatible DC
   If hScreenDC <> 0 Then MemoryDC_Create = CreateCompatibleDC(hScreenDC)
   ' Release the handle to the display screen
   ReleaseDC GetDesktopWindow, hScreenDC
End Function
' MemoryDC_Delete
'
' This function deletes any memory Device Context (DC) created by a call to such Win32 APIs as "CreateDC"
' or "CreateCompatibleDC".  Any DC created by such Win32 API calls *MUST* be deleted by calling this
' function or the "DeleteDC" Win32 API Function.  If you do not, you could cause a VERY LARGE memory
' leak in your program.
'
' Parameter:              Use:
' --------------------------------------------------
' hMemoryDC               Handle to the memory device context to delete
' hPrevBitmap             Handle to the previous bitmap of the device context.  This is returned to the
'                         DC before deleting the DC.
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
'
Public Function MemoryDC_Delete(ByRef hMemoryDC As Long, Optional ByRef hPrevBitmap As Long) As Boolean
   ' Make sure parameter(s) are valid
   If hMemoryDC = 0 Then Exit Function
   ' Put the previous bitmap back into the DC and delete the one that's in it now
   If hPrevBitmap <> 0 Then
      DeleteObject SelectObject(hMemoryDC, hPrevBitmap)
      hPrevBitmap = 0
   End If
   ' Delete the DC
   If DeleteDC(hMemoryDC) <> 0 Then
      MemoryDC_Delete = True
      hMemoryDC = 0
   End If
End Function
'
' LoadBitmapToDC
'
' This function loads the specified BITMAP into the specified Device Context (DC) and returns the handle
' to the picture that was previously in the specified DC.  This is important, because that picture
' should be reloaded back into the DC before deleting the DC.
'
' NOTE: You can call the "MemoryDC_Delete" function of this module and pass it the returned picture
' handle and it will properly restore the old picture, delete the new picture, then delete the DC.
'
' Parameter:              Use:
' --------------------------------------------------
' hDC                     Handle to the DC to load the picture into
' hPictureToLoad          Handle to the picture to load into the DC
' Return_hPrevPic         Optional. Returns the handle to the picture that was previously in the DC
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
'
Public Function LoadBitmapToDC(ByVal hdc As Long, ByVal hPictureToLoad As Long, Optional ByRef Return_hPrevPic As Long) As Boolean
   ' Make sure the values passed are valid
   If hdc = 0 Or hPictureToLoad = 0 Then Exit Function
   ' Select the specified picture into the specified DC
   Return_hPrevPic = SelectObject(hdc, hPictureToLoad)
   ' Return successful results
   LoadBitmapToDC = True
End Function
' RefreshHWND
'
' This function "refreshes" the specified window by invalidating it's contents.  This sends messages
' to the specified window that it needs to repaint itself.
'
' NOTE - If a picture is drawn on a VB Form, PictureBox, etc and that object's "AutoRedraw" property
' is set to TRUE, the picture will not immediately appear.  You must call this function, or the object's
' "Refresh" method to see the newly drawn picture on it.  However, if that object's "AutoRedraw" property
' is set to FALSE, the picture will immediately appear and any call to this function or that object's
' "Refresh" method will cause the picture to be erased.
'
' NOTE - Pictures are drawn to the "Image" property of a VB object, not the "Picture" property.  You
' can set the "Picture" property like this :  Set Picture1.Picture = Picture1.Image
'
' Parameter:              Use:
' --------------------------------------------------
' WindowHandle            Handle to the window or object to refresh
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
'
Public Function RefreshHWND(ByVal WindowHandle As Long) As Boolean
   On Error Resume Next
   ' If RedrawWindow(WindowHandle, ByVal 0, 0, RDW_INVALIDATE Or RDW_ALLCHILDREN) <> 0 Then RefreshHWND = True
   If RedrawWindow(WindowHandle, ByVal 0, 0, RDW_INVALIDATE) <> 0 Then RefreshHWND = True
End Function
' RenderIcon
'
' This function takes the specified icon and renders it to the specified Device Context.
'
' Parameter:              Use:
' --------------------------------------------------
' Dest_hDC                Specifies the Device Context to render the ICON onto
' hIcon                   Handle of the ICON to render
' Dest_X                  Optional. Specifies the X (Left) position to draw the ICON to on the DC
' Dest_Y                  Optional. Specifies the Y (Top) position to draw the ICON to on the DC
' Dest_Height             Optional. Specifies the height of the icon when rendered.  If not specified,
'                         the original icon height is used.
' Dest_Width              Optional. Specifies the width of the icon when rendered.  If not specified,
'                         the original icon width is used.
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
Public Function RenderIcon(ByVal Dest_hDC As Long, ByVal hIcon As Long, Optional ByVal Dest_X As Long, Optional ByVal Dest_Y As Long, Optional ByVal Dest_Height As Long, Optional ByVal Dest_Width As Long) As Boolean
   ' Make sure parameters passed are valid
   If Dest_hDC = 0 Or hIcon = 0 Then Exit Function
   If DrawIconEx(Dest_hDC, Dest_X, Dest_Y, hIcon, Dest_Width, Dest_Height, 0, 0, DI_NORMAL) <> 0 Then
      RenderIcon = True
   End If
End Function
' RenderIconGrayscale
'
' This function takes the specified icon and converts it to grayscale and then renders it to the
' specified Device Context.
'
' Parameter:              Use:
' --------------------------------------------------
' Dest_hDC                Specifies the Device Context to render the ICON onto
' hIcon                   Handle of the ICON to render
' Dest_X                  Optional. Specifies the X (Left) position to draw the ICON to on the DC
' Dest_Y                  Optional. Specifies the Y (Top) position to draw the ICON to on the DC
' Dest_Height             Optional. Specifies the height of the icon when rendered.  If not specified,
'                         the original icon height is used.
' Dest_Width              Optional. Specifies the width of the icon when rendered.  If not specified,
'                         the original icon width is used.
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
'
Public Function RenderIconGrayscale(ByVal Dest_hDC As Long, ByVal hIcon As Long, Optional ByVal Dest_X As Long, Optional ByVal Dest_Y As Long, Optional ByVal Dest_Height As Long, Optional ByVal Dest_Width As Long) As Boolean
   Dim hBMP_Mask  As Long
   Dim hBMP_Image As Long
   Dim hBMP_Prev  As Long
   Dim hIcon_Temp As Long
   Dim hDC_Temp   As Long
   ' Make sure parameters passed are valid
   If Dest_hDC = 0 Or hIcon = 0 Then Exit Function
   ' Extract the bitmaps from the icon
   If GetIconBitmaps(hIcon, hBMP_Mask, hBMP_Image) = False Then Exit Function
   ' Create a memory DC to work with
   hDC_Temp = MemoryDC_Create
   If hDC_Temp = 0 Then GoTo CleanUp
   ' Make the image bitmap gradient
   If RenderBitmapGrayscale(hDC_Temp, hBMP_Image, 0, 0, , , True) = False Then GoTo CleanUp
   ' Extract the gradient bitmap out of the DC
   SelectObject hDC_Temp, hBMP_Prev
   ' Take the newly gradient bitmap and make a gradient icon from it
   hIcon_Temp = CreateIconFromBMP(hBMP_Mask, hBMP_Image)
   If hIcon_Temp = 0 Then GoTo CleanUp
   ' Draw the newly created gradient icon onto the specified DC
   If DrawIconEx(Dest_hDC, Dest_X, Dest_Y, hIcon_Temp, Dest_Width, Dest_Height, 0, 0, DI_NORMAL) <> 0 Then
      RenderIconGrayscale = True
   End If
CleanUp:
   DestroyIcon hIcon_Temp: hIcon_Temp = 0
   DeleteDC hDC_Temp: hDC_Temp = 0
   DeleteObject hBMP_Mask: hBMP_Mask = 0
   DeleteObject hBMP_Image: hBMP_Image = 0
End Function
' RenderCursor
'
' This function takes the specified cursor and renders it to the specified Device Context.
'
' Parameter:              Use:
' --------------------------------------------------
' Dest_hDC                Specifies the Device Context to render the CURSOR onto
' hCursor                 Handle of the CURSOR to render
' Dest_X                  Optional. Specifies the X (Left) position to draw the CURSOR to on the DC
' Dest_Y                  Optional. Specifies the Y (Top) position to draw the CURSOR to on the DC
' Dest_Height             Optional. Specifies the height of the cursor when rendered.  If not specified,
'                         the original cursor height is used.
' Dest_Width              Optional. Specifies the width of the cursor when rendered.  If not specified,
'                         the original cursor width is used.
' AnimatedCursorFrame     Optioanl. If the cursor that is specified is an animated cursor (*.ANI) then
'                         you can render just one of the animation frames by specifying the index of the
'                         frame to render in this parameter.
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
'
Public Function RenderCursor(ByVal Dest_hDC As Long, ByVal hCursor As Long, Optional ByVal Dest_X As Long, Optional ByVal Dest_Y As Long, Optional ByVal Dest_Height As Long, Optional ByVal Dest_Width As Long, Optional ByVal AnimatedCursorFrame As Long) As Boolean
   ' Make sure parameters passed are valid
   If Dest_hDC = 0 Or hCursor = 0 Then Exit Function
   If DrawIconEx(Dest_hDC, Dest_X, Dest_Y, hCursor, Dest_Width, Dest_Height, AnimatedCursorFrame, 0, DI_NORMAL) <> 0 Then
      RenderCursor = True
   End If
End Function
' RenderCursorGrayscale
'
' This function takes the specified cursor and changes it to grayscale and then renders it to the
' specified Device Context.
'
' Parameter:              Use:
' --------------------------------------------------
' Dest_hDC                Specifies the Device Context to render the CURSOR onto
' hCursor                 Handle of the CURSOR to render
' Dest_X                  Optional. Specifies the X (Left) position to draw the CURSOR to on the DC
' Dest_Y                  Optional. Specifies the Y (Top) position to draw the CURSOR to on the DC
' Dest_Height             Optional. Specifies the height of the cursor when rendered.  If not specified,
'                         the original cursor height is used.
' Dest_Width              Optional. Specifies the width of the cursor when rendered.  If not specified,
'                         the original cursor width is used.
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
'
Public Function RenderCursorGrayscale(ByVal Dest_hDC As Long, ByVal hCursor As Long, Optional ByVal Dest_X As Long, Optional ByVal Dest_Y As Long, Optional ByVal Dest_Height As Long, Optional ByVal Dest_Width As Long) As Boolean
   Dim hBMP_Mask    As Long
   Dim hBMP_Image   As Long
   Dim hBMP_Prev    As Long
   Dim hCursor_Temp As Long
   Dim hDC_Temp     As Long
   ' Make sure parameters passed are valid
   If Dest_hDC = 0 Or hCursor = 0 Then Exit Function
   ' Extract the bitmaps from the cursor
   If GetIconBitmaps(hCursor, hBMP_Mask, hBMP_Image) = False Then Exit Function
   ' Create a memory DC to work with
   hDC_Temp = MemoryDC_Create
   If hDC_Temp = 0 Then GoTo CleanUp
   ' Make the image bitmap gradient
   If RenderBitmapGrayscale(hDC_Temp, hBMP_Image, 0, 0, , , True) = False Then GoTo CleanUp
   ' Extract the gradient bitmap out of the DC
   SelectObject hDC_Temp, hBMP_Prev
   ' Take the newly gradient bitmap and make a gradient cursor from it
   hCursor_Temp = CreateCursorFromBMP(hBMP_Mask, hBMP_Image)
   If hCursor_Temp = 0 Then GoTo CleanUp
   ' Draw the newly created gradient cursor onto the specified DC
   If DrawIconEx(Dest_hDC, Dest_X, Dest_Y, hCursor_Temp, Dest_Width, Dest_Height, 0, 0, DI_NORMAL) <> 0 Then
      RenderCursorGrayscale = True
   End If
CleanUp:
   DestroyIcon hCursor_Temp: hCursor_Temp = 0
   DeleteDC hDC_Temp: hDC_Temp = 0
   DeleteObject hBMP_Mask: hBMP_Mask = 0
   DeleteObject hBMP_Image: hBMP_Image = 0
End Function
' RenderBitmap
'
' This function takes the specified bitmap and renders it to the specified device context.  This is a
' more simplified version of RenderBitmapEx and has less overhead because it doesn't perform all of
' the calculations that RenderBitmapEx does.  It also doesn't do the strech functionality that
' RenderBitmapEx does.
'
' Parameter:              Use:
' --------------------------------------------------
' Dest_hDC                Specifies the Device Context to render the BITMAP onto
' hBitmap                 Handle of the BITMAP to render
' Dest_X                  Optional. Specifies the X (Left) position to draw the BITMAP to on the DC
' Dest_Y                  Optional. Specifies the Y (Top) position to draw the BITMAP to on the DC
' Srce_X                  Optional. Specifies the X (Left) position of the source picture that the
'                         picture should be drawn from
' Srce_Y                  Optional. Specifies the Y (Top) position to the source picture that the
'                         picture should be drawn from
' RasterOperation         Optional. Specifies the raster operation to perform on the image when
'                         BitBlt'ing it to the specified Device Context
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
'
'=============================================================================================================
Public Function RenderBitmap(ByVal Dest_hDC As Long, ByVal hBitmap As Long, Optional ByVal Dest_X As Long, Optional ByVal Dest_Y As Long, Optional ByVal Srce_X As Long, Optional ByVal Srce_Y As Long, Optional ByVal RasterOperation As RasterOperations = SRCCOPY) As Boolean
   Dim TempBMP    As BITMAP
   Dim hDC_Temp   As Long
   Dim hDC_Screen As Long
   Dim hBMP_Prev  As Long
   ' Make sure parameters passed are valid
   If Dest_hDC = 0 Or hBitmap = 0 Then Exit Function
   ' Get the information about the bitmap (and make sure it's a valid BITMAP
   If GetObjectAPI(hBitmap, Len(TempBMP), TempBMP) = 0 Then Exit Function
   ' Create a memory DC to use for the render operation
   hDC_Screen = GetDC(GetDesktopWindow)
   If hDC_Screen = 0 Then Exit Function
   hDC_Temp = CreateCompatibleDC(hDC_Screen)
   If hDC_Temp = 0 Then GoTo CleanUp
   ' Select the specified bitmap into the memory DC just created
   hBMP_Prev = SelectObject(hDC_Temp, hBitmap)
   ' Render the bitmap onto the specified hDC
   If BitBlt(Dest_hDC, Dest_X, Dest_Y, TempBMP.bmWidth, TempBMP.bmHeight, hDC_Temp, Srce_X, Srce_Y, RasterOperation) <> 0 Then
      RenderBitmap = True
   End If
CleanUp:
   ReleaseDC GetDesktopWindow, hDC_Screen: hDC_Screen = 0
   SelectObject hDC_Temp, hBMP_Prev
   DeleteDC hDC_Temp: hDC_Temp = 0
End Function
'=============================================================================================================
'
' RenderBitmapEx
'
' This function takes the handle to a Picture (ie - PictureBox1.Picture.Handle) or the handle to a Device
' Context (DC) and renders (or paints) it onto the specified output DC (ie - PictureBox2.hDC).  This
' function also gives you the ability to stretch the picture to a specified height/width before rendering
' it to the output DC.  As an aditional option, you can refresh a specified window that is related to
' the specified DC (ie - PictureBox2.hWnd).
'
' NOTE : All measurements for this function are expected to be PIXELS.  The default measurement of
'        pictures in Visual Basic is HIMETRIC.  Convert the dimentions to pixels if need be using the
'        "ConvPicDimentions" function before calling this function, or use the "GetBitmapInfo" function
'        to get the height and width in pixels.
'
' NOTE : If you specify a form as the source hDC and do NOT specify a picture, you must specify the
'        picture's height/width.  Otherwise, the image rendered will have the same height and width as
'        your current screen resolution (800x600, 1024x768, etc)
'
' NOTE : This function is the most efficient if you specify a handle to a DC in the Source_hDC parameter
'        that contains the picture specified in the hBitmap parameter as well as the height and width
'        of the picture specified.  That way no DC's have to be created or queried.
'
' Parameter:              Use:
' --------------------------------------------------
' Dest_hDC                The DC that the image is to be drawn onto
' Source_hDC              Optional. If this parameter is specified and the hBitmap parameter is not,
'                         this function takes the currently selected bitmap from the DC and uses it
' hBitmap                 Optional. If this parameter is specified, a memory DC is created and the
'                         picture is selected into it to be used to paint the picture to the specified
'                         DC.  If this parameter is NOT specified, there must be a valid DC in the
'                         Source_DC parameter containing a valid picture to use.
' Dest_X                  Optional. Specifies X (Left) coordinate of where the picture should be rendered
'                         on the specified Dest_hDC (pixels).
' Dest_Y                  Optional. Specifies Y (Top) coordinate of where the picture should be rendered
'                         on the specified Dest_hDC (pixels).
' Srce_X                  Optional. Specifies X (Left) coordinate of where the picture should be taken
'                         from on the specified Source_hDC (pixels).
' Srce_Y                  Optional. Specifies Y (Top) coordinate of where the picture should be taken
'                         from on the specified Source_hDC (pixels).
' PicHeight               Optional. This specifies the height (in pixels) of the picture to be rendered.
'                         If this parameter is NOT specified, the function attempts to find out the
'                         height of the picture and use that.
' PicWidth                Optional. This specifies the width (in pixels) of the picture to be rendered.
'                         If this parameter is NOT specified, the function attempts to find out the
'                         width of the picture and use that.
' RasterOperation         Optional. This specifies the raster operation to be used in painting the
'                         picture.  The default is SRCCOPY (which just copies the picture).
' StretchPicture          Optional. If this parameter is set to TRUE, the picture is stretched to the
'                         size specified in the StretchHeight and StretchWidth parameters.
' StretchHeight           Optional. If the StretchPicture parameter is set to TRUE, this parameter
'                         specifies the height to stretch the picture to.
' StretchWidth            Optional. If the StretchPicture parameter is set to TRUE, this parameter
'                         specifies the width to stretch the picture to.
' RefreshWindow           Optional. If this parameter is set to TRUE, the function refreshes the window
'                         specified in the RefreshHandle parameter (which should be the handle of the
'                         window associated with the DC specified in the Dest_hDC parameter.
' RefreshHandle           Optional. If the RefreshWindow parameter is set to TRUE, this parameter
'                         specifies the handle of the window to refresh.
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
'
Public Function RenderBitmapEx(ByVal Dest_hDC As Long, Optional ByVal Source_hDC As Long, Optional ByVal hBitmap As Long, Optional ByVal Dest_X As Long, Optional ByVal Dest_Y As Long, Optional ByVal Srce_X As Long, Optional ByVal Srce_Y As Long, Optional ByVal PicHeight As Long, Optional ByVal PicWidth As Long, Optional ByVal RasterOperation As RasterOperations = SRCCOPY, Optional ByVal StretchPicture As Boolean = False, Optional ByVal StretchHeight As Long, Optional ByVal StretchWidth As Long, Optional ByVal RefreshWindow As Boolean = False, Optional ByVal RefreshHandle As Long) As Boolean
   On Error Resume Next
   Dim ScrHWND     As Long
   Dim ScrHDC      As Long
   Dim hMemoryDC   As Long
   Dim BMP         As BITMAP
   Dim bDelMemDC   As Boolean
   Dim hOldBitmap  As Long
   Dim PrevStrMode As Long
   ' Make sure parameters passed are valid
   If Dest_hDC = 0 Then
      Exit Function
   ElseIf Source_hDC = 0 And hBitmap = 0 Then
      Exit Function
   End If
   ' If no source DC was specified, but a picture was specified, create a DC to use
   If Source_hDC = 0 And hBitmap <> 0 Then
      bDelMemDC = True
      ScrHWND = GetDesktopWindow
      ScrHDC = GetDC(ScrHWND)
      hMemoryDC = CreateCompatibleDC(ScrHDC)
      ReleaseDC ScrHWND, ScrHDC
      If hMemoryDC = 0 Then GoTo CleanUp
      hOldBitmap = SelectObject(hMemoryDC, hBitmap)
      ' If a source DC was specified and no picture was specified, make sure the source hDC has a picture
   ElseIf Source_hDC <> 0 And hBitmap = 0 Then
      bDelMemDC = False
      hMemoryDC = Source_hDC
      hBitmap = GetBitmapFromDC(Source_hDC)
      If hBitmap = 0 Then
         Exit Function
      End If
      ' If a source DC AND a picture was specified, use them both
   ElseIf Source_hDC <> 0 And hBitmap <> 0 Then
      bDelMemDC = False
      hMemoryDC = Source_hDC
   End If
   ' If the user didn't specify a Height / Width, get them from the picture
   If PicHeight = 0 Or PicWidth = 0 Then
      If GetObjectAPI(hBitmap, Len(BMP), BMP) = 0 Then GoTo CleanUp
      If PicHeight = 0 Then PicHeight = BMP.bmHeight
      If PicWidth = 0 Then PicWidth = BMP.bmWidth
   End If
   ' Check if the user wants to stretch the picture
   If StretchPicture = True Then
      If StretchHeight = PicHeight And StretchWidth = PicWidth Then
         StretchPicture = False
      ElseIf StretchHeight = 0 Or StretchWidth = 0 Then
         StretchPicture = False
      End If
   End If
   ' Render the picture onto the specified DC
   If StretchPicture = False Then
      If BitBlt(Dest_hDC, Dest_X, Dest_Y, PicWidth, PicHeight, hMemoryDC, Srce_X, Srce_Y, RasterOperation) <> 0 Then
         RenderBitmapEx = True
      End If
   Else
      PrevStrMode = SetStretchBltMode(Dest_hDC, STRETCH_HALFTONE) ' This DRAMATICALLY improves the quality of the following stretch operation
      If SetBrushOrgEx(Dest_hDC, 0, 0, ByVal 0) = 0 Then GoTo CleanUp
      If StretchBlt(Dest_hDC, Dest_X, Dest_Y, StretchWidth, StretchHeight, hMemoryDC, Srce_X, Srce_Y, PicWidth, PicHeight, RasterOperation) <> 0 Then
         RenderBitmapEx = True
      End If
      SetStretchBltMode Dest_hDC, PrevStrMode
   End If
   ' Refresh the DC to show the picture just drawn on it
   If RenderBitmapEx = True And RefreshWindow = True And RefreshHandle <> 0 Then
      RefreshHWND RefreshHandle
   End If
CleanUp:
   If bDelMemDC = True And hMemoryDC <> 0 Then
      SelectObject hMemoryDC, hOldBitmap
      DeleteDC hMemoryDC: hMemoryDC = 0
   End If
End Function
' RenderBitmapGrayscale
'
' This function takes the specified bitmap and converts it to grayscale before rendering it to the
' specified Device Context.
'
' Parameter:              Use:
' --------------------------------------------------
' Dest_hDC                Specifies the Device Context to render the BITMAP onto
' hBitmap                 Handle of the BITMAP to render
' Dest_X                  Optional. Specifies the X (Left) position to draw the BITMAP to on the DC
' Dest_Y                  Optional. Specifies the Y (Top) position to draw the BITMAP to on the DC
' Srce_X                  Optional. Specifies the X (Left) position of the source picture that the
'                         picture should be drawn from
' Srce_Y                  Optional. Specifies the Y (Top) position to the source picture that the
'                         picture should be drawn from
' AlterOriginalPic        Optional. If TRUE, the original picture that is passed to this function is
'                         changed to grayscale before it is rendered to the specified Device Context.
'                         If FALSE, a copy is made of the picture, changed to grayscale, and then
'                         rendered to the specified DC.
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
'
Public Function RenderBitmapGrayscale(ByVal Dest_hDC As Long, ByVal hBitmap As Long, Optional ByVal Dest_X As Long, Optional ByVal Dest_Y As Long, Optional ByVal Srce_X As Long, Optional ByVal Srce_Y As Long, Optional ByVal AlterOriginalPic As Boolean = False) As Boolean
   Dim TempBITMAP  As BITMAP
   Dim hScreen     As Long
   Dim hDC_Temp    As Long
   Dim hBMP_Prev   As Long
   Dim MyCounterX  As Long
   Dim MyCounterY  As Long
   Dim NewColor    As Long
   Dim hNewPicture As Long
   Dim DeletePic   As Boolean
   ' Make sure parameters passed are valid
   If Dest_hDC = 0 Or hBitmap = 0 Then Exit Function
   ' Get the handle to the screen DC
   hScreen = GetDC(GetDesktopWindow)
   If hScreen = 0 Then Exit Function
   ' Create a memory DC to work with the picture
   hDC_Temp = CreateCompatibleDC(hScreen)
   If hDC_Temp = 0 Then GoTo CleanUp
   ' If the user specifies NOT to alter the original, then make a copy of it to use
   If AlterOriginalPic = False Then
      DeletePic = True
      If CopyPicture(hBitmap, hNewPicture) = False Then GoTo CleanUp
   Else
      DeletePic = False
      hNewPicture = hBitmap
   End If
   ' Select the bitmap into the DC
   hBMP_Prev = SelectObject(hDC_Temp, hNewPicture)
   ' Get the height / width of the bitmap in pixels
   If GetObjectAPI(hNewPicture, Len(TempBITMAP), TempBITMAP) = 0 Then GoTo CleanUp
   If TempBITMAP.bmHeight <= 0 Or TempBITMAP.bmWidth <= 0 Then GoTo CleanUp
   ' Loop through each pixel and conver it to it's grayscale equivelant
   For MyCounterX = 0 To TempBITMAP.bmWidth - 1
      For MyCounterY = 0 To TempBITMAP.bmHeight - 1
         NewColor = GetPixel(hDC_Temp, MyCounterX, MyCounterY)
         If NewColor <> -1 Then
            Select Case NewColor
               ' If the color is already a grey shade, no need to convert it
            Case vbBlack, vbWhite, &H101010, &H202020, &H303030, &H404040, &H505050, &H606060, &H707070, &H808080, &HA0A0A0, &HB0B0B0, &HC0C0C0, &HD0D0D0, &HE0E0E0, &HF0F0F0
               NewColor = NewColor
            Case Else
               NewColor = GrayScale(NewColor)
            End Select
            SetPixel hDC_Temp, MyCounterX, MyCounterY, NewColor
         End If
      Next MyCounterY
   Next MyCounterX
   ' Display the picture on the specified hDC
   BitBlt Dest_hDC, Dest_X, Dest_Y, TempBITMAP.bmWidth, TempBITMAP.bmHeight, hDC_Temp, Srce_X, Srce_Y, vbSrcCopy
   RenderBitmapGrayscale = True
CleanUp:
   ReleaseDC GetDesktopWindow, hScreen: hScreen = 0
   SelectObject hDC_Temp, hBMP_Prev
   DeleteDC hDC_Temp: hDC_Temp = 0
   If DeletePic = True Then
      DeleteObject hNewPicture
      hNewPicture = 0
   End If
End Function
' RenderBitmapTransparentGS
'
' This function takes the specified BITMAP and first changes it to grayscale, then renders it onto the
' specified Device Context (DC) but does NOT render the specified transparent color (thus making it
' see-thru, or transparent)
'
' Parameter:              Use:
' --------------------------------------------------
' Dest_hDC                Specifies the Device Context to render the BITMAP onto
' hBitmap                 Handle of the BITMAP to render
' TransparentColor        Specifies the transparent color (color NOT to render)
' Dest_X                  Optional. Specifies the X (Left) position to draw the BITMAP to on the DC
' Dest_Y                  Optional. Specifies the Y (Top) position to draw the BITMAP to on the DC
' Srce_X                  Optional. Specifies the X (Left) position of the source picture that the
'                         picture should be drawn from
' Srce_Y                  Optional. Specifies the Y (Top) position to the source picture that the
'                         picture should be drawn from
' hDC_Background          Optional. If specified, the BITMAP contained within this Device Context will
'                         be used to draw the background
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
'
Public Function RenderBitmapTransparentGS(ByVal Dest_hDC As Long, ByVal hBitmap As Long, ByVal TransparentColor As Long, Optional ByVal Dest_X As Long, Optional ByVal Dest_Y As Long, Optional ByVal Srce_X As Long, Optional ByVal Srce_Y As Long, Optional ByVal hDC_Background As Long) As Boolean
   Dim TempBITMAP As BITMAP
   Dim hDC_Temp   As Long
   Dim hBMP_Copy  As Long
   Dim hBMP_Prev  As Long
   Dim MyCounterX As Long
   Dim MyCounterY As Long
   Dim NewColor   As Long
   ' Make sure parameters passed are valid
   If Dest_hDC = 0 Or hBitmap = 0 Then Exit Function
   ' Make sure the transparent color is a valid Win32 color
   TransparentColor = TranslateColor(TransparentColor)
   If TransparentColor = -1 Then Exit Function
   ' Create a DC to work with
   hDC_Temp = MemoryDC_Create
   If hDC_Temp = 0 Then Exit Function
   ' Create a copy of the orginal picture so we don't alter the original
   If CopyPicture(hBitmap, hBMP_Copy) = False Then Exit Function
   If hBMP_Copy = 0 Then GoTo CleanUp
   ' Get the information about the bitmap
   If GetObjectAPI(hBMP_Copy, Len(TempBITMAP), TempBITMAP) = 0 Then GoTo CleanUp
   ' Select the picture INTO the DC in order to work with it
   hBMP_Prev = SelectObject(hDC_Temp, hBMP_Copy)
   ' Loop through each pixel and conver it to it's grayscale equivelant
   For MyCounterX = 0 To TempBITMAP.bmWidth - 1
      For MyCounterY = 0 To TempBITMAP.bmHeight - 1
         NewColor = GetPixel(hDC_Temp, MyCounterX, MyCounterY)
         If NewColor <> -1 Then
            Select Case NewColor
               ' If the color is already a grey shade, no need to convert it
            Case TransparentColor, vbBlack, vbWhite, &H101010, &H202020, &H303030, &H404040, &H505050, &H606060, &H707070, &H808080, &HA0A0A0, &HB0B0B0, &HC0C0C0, &HD0D0D0, &HE0E0E0, &HF0F0F0
               NewColor = NewColor
            Case Else
               NewColor = GrayScale(NewColor)
            End Select
            SetPixel hDC_Temp, MyCounterX, MyCounterY, NewColor
         End If
      Next MyCounterY
   Next MyCounterX
   ' Select the picture OUT OF the DC to paint it
   SelectObject hDC_Temp, hBMP_Prev
   ' Render the grayscale bitmap transparently
   If RenderBitmapTransparent(Dest_hDC, hBMP_Copy, TransparentColor, Dest_X, Dest_Y, Srce_X, Srce_Y, hDC_Background) = True Then
      RenderBitmapTransparentGS = True
   End If
CleanUp:
   DeleteDC hDC_Temp: hDC_Temp = 0
   DeleteObject hBMP_Copy: hBMP_Copy = 0
End Function
' RenderBitmapTransparent
'
' This function takes the specified BITMAP and renders it onto the specified Device Context (DC) but
' does NOT render the specified transparent color (thus making it see-thru, or transparent)
'
' NOTE : This function is a modified version of a sample function taken from the VB 5.0 CD-ROM
'        (TOOLS\UNSUPPRT\SSAVER\PAINTSUP.BAS)
'
' Parameter:              Use:
' --------------------------------------------------
' Dest_hDC                Specifies the Device Context to render the BITMAP onto
' hBitmap                 Handle of the BITMAP to render
' TransparentColor        Specifies the transparent color (color NOT to render)
' Dest_X                  Optional. Specifies the X (Left) position to draw the BITMAP to on the DC
' Dest_Y                  Optional. Specifies the Y (Top) position to draw the BITMAP to on the DC
' Srce_X                  Optional. Specifies the X (Left) position of the source picture that the
'                         picture should be drawn from
' Srce_Y                  Optional. Specifies the Y (Top) position to the source picture that the
'                         picture should be drawn from
' hDC_Background          Optional. If specified, the BITMAP contained within this Device Context will
'                         be used to draw the background
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
'
Public Function RenderBitmapTransparent(ByVal Dest_hDC As Long, ByVal hBitmap As Long, ByVal TransparentColor As Long, Optional ByVal Dest_X As Long, Optional ByVal Dest_Y As Long, Optional ByVal Srce_X As Long, Optional ByVal Srce_Y As Long, Optional ByVal hDC_Background As Long) As Boolean
   Dim TempBITMAP     As BITMAP
   Dim PreviousColor  As Long  'COLORREF
   Dim hBMP_AndBack   As Long  'HBITMAP
   Dim hBMP_BackOld   As Long  'HBITMAP
   Dim hBMP_AndObject As Long  'HBITMAP
   Dim hBMP_ObjectOld As Long  'HBITMAP
   Dim hBMP_AndMem    As Long  'HBITMAP
   Dim hBMP_MemOld    As Long  'HBITMAP
   Dim hBMP_Save      As Long  'HBITMAP
   Dim hBMP_SaveOld   As Long  'HBITMAP
   Dim hDC_Mem        As Long  'HDC
   Dim hDC_Back       As Long  'HDC
   Dim hDC_Object     As Long  'HDC
   Dim hDC_Temp       As Long  'HDC
   Dim hDC_Save       As Long  'HDC
   Dim PicWidth       As Long
   Dim PicHeight      As Long
   ' Make sure parameters passed are valid
   If Dest_hDC = 0 Or hBitmap = 0 Then Exit Function
   ' Make sure the transparent color is a valid Win32 color
   TransparentColor = TranslateColor(TransparentColor)
   If TransparentColor = -1 Then Exit Function
   ' Create the DC to work from and get info on the bitmap
   hDC_Temp = CreateCompatibleDC(Dest_hDC)
   SelectObject hDC_Temp, hBitmap
   GetObjectAPI hBitmap, Len(TempBITMAP), TempBITMAP
   PicWidth = TempBITMAP.bmWidth
   PicHeight = TempBITMAP.bmHeight
   ' Create some DCs to hold temporary data
   hDC_Back = CreateCompatibleDC(Dest_hDC)
   hDC_Object = CreateCompatibleDC(Dest_hDC)
   hDC_Mem = CreateCompatibleDC(Dest_hDC)
   hDC_Save = CreateCompatibleDC(Dest_hDC)
   ' Monochrome DC
   hBMP_AndBack = CreateBitmap(PicWidth, PicHeight, 1, 1, 0)
   ' Monochrome DC
   hBMP_AndObject = CreateBitmap(PicWidth, PicHeight, 1, 1, 0)
   ' Compatible DC's
   hBMP_AndMem = CreateCompatibleBitmap(Dest_hDC, PicWidth, PicHeight)
   hBMP_Save = CreateCompatibleBitmap(Dest_hDC, PicWidth, PicHeight)
   ' Each DC must select a bitmap object to store pixel data.
   hBMP_BackOld = SelectObject(hDC_Back, hBMP_AndBack)
   hBMP_ObjectOld = SelectObject(hDC_Object, hBMP_AndObject)
   hBMP_MemOld = SelectObject(hDC_Mem, hBMP_AndMem)
   hBMP_SaveOld = SelectObject(hDC_Save, hBMP_Save)
   ' Set proper mapping mode.
   SetMapMode hDC_Temp, GetMapMode(Dest_hDC)
   ' Save the bitmap sent here, because it will be overwritten
   BitBlt hDC_Save, 0, 0, PicWidth, PicHeight, hDC_Temp, Srce_X, Srce_Y, vbSrcCopy
   ' Set the background color of the source DC to the color contained in the parts of the bitmap that should be transparent
   PreviousColor = SetBkColor(hDC_Temp, TransparentColor)
   ' Create the object mask for the bitmap by performaing a BitBlt from the source bitmap to a monochrome bitmap.
   BitBlt hDC_Object, 0, 0, PicWidth, PicHeight, hDC_Temp, Srce_X, Srce_Y, vbSrcCopy
   ' Set the background color of the source DC back to the original color
   SetBkColor hDC_Temp, PreviousColor
   ' Create the inverse of the object mask.
   BitBlt hDC_Back, 0, 0, PicWidth, PicHeight, hDC_Object, 0, 0, vbNotSrcCopy
   ' Copy the background of the main DC to the destination
   If hDC_Background <> 0 Then
      BitBlt hDC_Mem, 0, 0, PicWidth, PicHeight, hDC_Background, Dest_X, Dest_Y, vbSrcCopy
   Else
      BitBlt hDC_Mem, 0, 0, PicWidth, PicHeight, Dest_hDC, Dest_X, Dest_Y, vbSrcCopy
   End If
   ' Mask out the places where the bitmap will be placed
   BitBlt hDC_Mem, 0, 0, PicWidth, PicHeight, hDC_Object, 0, 0, vbSrcAnd
   ' Mask out the transparent colored pixels on the bitmap
   BitBlt hDC_Temp, Srce_X, Srce_Y, PicWidth, PicHeight, hDC_Back, 0, 0, vbSrcAnd
   ' XOR the bitmap with the background on the destination DC
   BitBlt hDC_Mem, 0, 0, PicWidth, PicHeight, hDC_Temp, Srce_X, Srce_Y, vbSrcPaint
   ' Copy the destination to the screen
   BitBlt Dest_hDC, Dest_X, Dest_Y, PicWidth, PicHeight, hDC_Mem, 0, 0, vbSrcCopy
   ' Place the original bitmap back into the bitmap sent here
   BitBlt hDC_Temp, Srce_X, Srce_Y, PicWidth, PicHeight, hDC_Save, 0, 0, vbSrcCopy
   ' Delete memory bitmaps
   DeleteObject SelectObject(hDC_Back, hBMP_BackOld): hBMP_AndBack = 0
   DeleteObject SelectObject(hDC_Object, hBMP_ObjectOld): hBMP_AndObject = 0
   DeleteObject SelectObject(hDC_Mem, hBMP_MemOld): hBMP_AndMem = 0
   DeleteObject SelectObject(hDC_Save, hBMP_SaveOld): hBMP_Save = 0
   ' Delete memory DC's
   DeleteDC hDC_Back: hDC_Back = 0
   DeleteDC hDC_Mem: hDC_Mem = 0
   DeleteDC hDC_Object: hDC_Object = 0
   DeleteDC hDC_Temp: hDC_Temp = 0
   DeleteDC hDC_Save: hDC_Save = 0
   RenderBitmapTransparent = True
End Function
' TileBitmap
'
' This function makes it easy to tile the specified picture onto any Device Context (DC).
'
' Parameter:              Use:
' --------------------------------------------------
' Dest_hDC                Specifies the Device Context to tile to (must be pre-initialized and the same
'                         size as is specified in the "Dest_Width" and "Dest_Height" perameters - You
'                         can size a memory DC by creating a BITMAP in memory to the size you want the
'                         DC to be by calling the "CreateCompatibleBitmap" API, then SelectObject the
'                         BITMAP into the DC).
' hBitmap                 Specifies the handle to the image to tile
' Dest_Width              Specifies the width in pixels of the DC specified in the "Dest_hDC" parameter
' Dest_Height             Specifies the height in pixels of the DC specified in the "Dest_hDC" parameter
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
'
Public Function TileBitmap(ByVal Dest_hDC As Long, ByVal hBitmap As Long, ByVal Dest_Width As Long, ByVal Dest_Height As Long) As Boolean
   Dim CurrentX   As Long
   Dim CurrentY   As Long
   Dim TempBITMAP As BITMAP
   ' Make sure the parameters passed are VALID
   If hBitmap = 0 Or Dest_hDC = 0 Or Dest_Width <= 0 Or Dest_Height <= 0 Then Exit Function
   ' Get the dimentions of the specified bitmap (this also verifies the image is truely a BITMAP)
   If GetObjectAPI(hBitmap, Len(TempBITMAP), TempBITMAP) = 0 Then Exit Function
   ' Line by line, row by row, tile the picture into the specified DC
   While CurrentX < Dest_Width
      While CurrentY < Dest_Height
         If RenderBitmap(Dest_hDC, hBitmap, CurrentX, CurrentY, 0, 0, SRCCOPY) = False Then Exit Function
         CurrentY = CurrentY + TempBITMAP.bmHeight
      Wend
      CurrentY = 0
      CurrentX = CurrentX + TempBITMAP.bmWidth
   Wend
   TileBitmap = True
End Function
' Function that converts automation colors such as "vbButtonFace" to standard
' color such as "12632256".  It is safest to pass all colors through this
' function to make sure that if a user passes a color like "Me.BackColor" and
' the BackColor is vbButtonFace, it won't mess up any of the API's that are
' expecting a normal color value.
Public Function TranslateColor(ByVal oClr As Long, Optional ByVal hPal As Long = 0) As Long
   On Error Resume Next
   If OleTranslateColor(oClr, hPal, TranslateColor) <> 0 Then TranslateColor = -1
End Function
' Takes a color value and converts it to an equivalent grayscale value
Private Function GrayScale(ByVal ColorToConvert As Long, Optional ByVal ExemptColor As Long = -1) As Long
   If ExemptColor = -1 Then
      ColorToConvert = 0.33 * (ColorToConvert Mod 256) + 0.59 * ((ColorToConvert \ 256) Mod 256) + 0.11 * ((ColorToConvert \ 65536) Mod 256)
      GrayScale = RGB(ColorToConvert, ColorToConvert, ColorToConvert)
      Exit Function
   Else
      If ColorToConvert <> ExemptColor Then
         ColorToConvert = 0.33 * (ColorToConvert Mod 256) + 0.59 * ((ColorToConvert \ 256) Mod 256) + 0.11 * ((ColorToConvert \ 65536) Mod 256)
         GrayScale = RGB(ColorToConvert, ColorToConvert, ColorToConvert)
         Exit Function
      End If
   End If
   GrayScale = ColorToConvert
End Function
