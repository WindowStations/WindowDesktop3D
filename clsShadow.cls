VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const ShadeWidth As Single = 256
Private Const sz         As Long = 4096
Private VS               As Direct3DVertexShader9
Private ps               As Direct3DPixelShader9
Private RTSurf           As Direct3DSurface9
Private RTZB             As Direct3DSurface9
Private devSurf          As Direct3DSurface9
Private devZB            As Direct3DSurface9
Private Lite1Pos         As D3DVECTOR
Private Center           As D3DVECTOR
Friend Sub Initialize()
    On Error Resume Next
    Center = Vec3(128.0512, 32, 128)
    Vec3Scale Lite1Pos, SunLightPos, ShadeWidth / 2
    Vec3Add Lite1Pos, Center, Lite1Pos
    MatrixLookAtLH mViewShad, Lite1Pos, Center, Vec3(0, 1, 0)
    MatrixOrthoLH mProjShad, ShadeWidth, ShadeWidth, 0.1, ShadeWidth
    VSCreate
    PSCreate
    CreateSurf
End Sub
Friend Sub DrawBegin()
    On Error Resume Next
    Dim Mtrx As D3DMATRIX
    Dev.SetRenderTarget 0, RTSurf
    Dev.SetDepthStencilSurface RTZB
    Dev.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE
    Dev.SetRenderState D3DRS_ALPHABLENDENABLE, D3D_FALSE
    Dev.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    Dev.SetVertexShader VS
    Dev.SetPixelShader ps
    Dev.Clear D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &HFFFFFF, 1, 0
End Sub
Friend Sub DrawEnd()
    On Error Resume Next
    Dev.SetRenderTarget 0, devSurf
    Dev.SetDepthStencilSurface devZB
End Sub
Private Sub VSCreate()
    On Error Resume Next
    Set VS = Dev.CreateVertexShaderFromFile(App.Path & "\3DEngine\ShaderCompiler\Out\Shad.vsh.shader")
End Sub
Private Sub PSCreate()
    On Error Resume Next
    Set ps = Dev.CreatePixelShaderFromFile(App.Path & "\3DEngine\ShaderCompiler\Out\Shad.psh.shader")
End Sub
Friend Sub CreateSurf()
    On Error Resume Next
    Set devSurf = Dev.GetRenderTarget(0)
    Set devZB = Dev.GetDepthStencilSurface
    Set RTTex = Dev.CreateTexture(sz, sz, 1, D3DUSAGE_RENDERTARGET, D3DFMT_R32F, D3DPOOL_DEFAULT)
    Set RTSurf = RTTex.GetSurfaceLevel(0)
    Set RTZB = Dev.CreateDepthStencilSurface(sz, sz, D3DFMT_D24X8, D3DMULTISAMPLE_NONE, 0, D3D_FALSE)
End Sub
Friend Sub Terminate()
    On Error Resume Next
    Set VS = Nothing
    Set ps = Nothing
    KillSurf
End Sub
Friend Sub KillSurf()
    On Error Resume Next
    Set RTTex = Nothing
    Set RTSurf = Nothing
    Set RTZB = Nothing
    Set devZB = Nothing
    Set devSurf = Nothing
End Sub
