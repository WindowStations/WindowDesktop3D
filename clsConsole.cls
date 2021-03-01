VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
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
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Long, ByRef lpbi As BITMAPINFO, ByVal wUsage As Long) As Long
Private bi32BitInfo   As BITMAPINFO
Private TEX1          As Direct3DTexture9 'center scope texture
Private TexMSG        As Direct3DTexture9
Private SurfMSG       As Direct3DSurface9
Private cBuf()        As Long
Private MSGTimer      As Double
Private MSGDelay      As Long
Private MSGFade       As Double
Friend Sub Initialize()
    On Error Resume Next
    ReDim cBuf(1000 * 32 - 1) 'Set center scope and on screen message display
    Set TEX1 = CreateTextureFromFileEx(Dev, App.Path & "\3DEngine\tex\1.png", 0, 0, 0, D3DUSAGE_NONE, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_BOX, D3DX_FILTER_LINEAR, 0)
    Set TexMSG = Dev.CreateTexture(1000, 32, 1, D3DUSAGE_NONE, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED)
    Set SurfMSG = TexMSG.GetSurfaceLevel(0)
    frmMain.picMSG.FontSize = 0.7 * Screen.TwipsPerPixelX
End Sub
Friend Sub Draw()
    On Error Resume Next
    Dev.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    Dev.SetRenderState D3DRS_ZENABLE, D3DZB_FALSE
    Dev.SetRenderState D3DRS_ALPHABLENDENABLE, D3D_TRUE
    Dev.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    Dev.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    If MSGRender = True Then 'Render only if idle in a dead zone on selected object
        frmMain.spriteClass.Draw TexMSG, BBWidth + 32 - MSGTimer, (BBHeight / 2) + displayPosition - 500, 2000, 64, MSGFade * 0.001 '  frmMain.spriteClass.Draw TexMSG, (BBWidth \ 2) + 400, (BBHeight \ 2), 2500, 128, MSGTimer * 0.001
    End If
    frmMain.spriteClass.Draw TEX1, BBWidth / 2, BBHeight / 2, 64, 64 ' draw center scope
End Sub
Friend Sub Tick()
    On Error Resume Next
    If MSGDelay > 0 Then
        MSGDelay = MSGDelay - 1
    Else
        MSGTimer = MSGTimer + (displaySpeed / 3) '0.5
        MSGFade = MSGFade - 1
    End If
End Sub
Friend Sub Display(Msg As String)
    On Error Resume Next
    If displayHide <> 0 Then Exit Sub
    
  
    
    Dim r As D3DRECT
    r.left = 0
    r.right = 1000
    r.top = 0
    r.bottom = 32
    With bi32BitInfo.bmiHeader
        .biBitCount = 32
        .biPlanes = 1
        .biSize = Len(bi32BitInfo.bmiHeader)
        .biWidth = 1000
        .biHeight = -32
        .biSizeImage = 1000& * 32 * 4
    End With
    frmMain.picMSG.Cls
    '    frmMain.picMSG.Line (0, 5)-(1000, 32), &H0&, BF
    '    frmMain.picMSG.ScaleLeft = 0
    '    frmMain.picMSG.top = 2
    '    frmMain.picMSG.left = 2
    '    frmMain.picMSG.Width = 256
    '    frmMain.picMSG.Height = 32
    '    frmMain.picMSG.CurrentX = 0
    frmMain.picMSG.CurrentY = 5
    '    frmMain.picMSG.ScaleWidth = 256
    '    frmMain.picMSG.ScaleHeight = 32
    frmMain.picMSG.Print Msg;
    GetDIBits frmMain.picMSG.hdc, frmMain.picMSG.Image.Handle, 0, 24, cBuf(0), bi32BitInfo, 0
    SurfMSG.SetData VarPtr(r), 1000 * 32 * 4, VarPtr(cBuf(0))
    MSGTimer = 0
    MSGDelay = displayDelay
    MSGFade = displayFade
    MSGRender = True
End Sub
Friend Sub Terminate()
    Set SurfMSG = Nothing
    Set TexMSG = Nothing
    Set TEX1 = Nothing
End Sub
