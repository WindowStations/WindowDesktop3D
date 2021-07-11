VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Direct3DDevice9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub dev_Release Lib "dx9_vb" (ByVal pDev As Long)
Private Declare Sub dev_Clear Lib "dx9_vb" (ByVal pDev As Long, ByVal flags As CONST_D3DCLEARFLAGS, ByVal Color As Long, ByVal z As Single, ByVal Stencil As Long)
Private Declare Function dev_BeginScene Lib "dx9_vb" (ByVal pDev As Long) As Long
Private Declare Sub dev_EndScene Lib "dx9_vb" (ByVal pDev As Long)
Private Declare Sub dev_Present Lib "dx9_vb" (ByVal pDev As Long)
Private Declare Sub dev_SetStreamSource Lib "dx9_vb" (ByVal pDev As Long, ByVal StreamNumber As Long, ByVal pVB As Long, ByVal Offset As Long, ByVal Stride As Long)
Private Declare Sub dev_SetFVF Lib "dx9_vb" (ByVal pDev As Long, ByVal FVF As CONST_D3DFVF)
Private Declare Sub dev_SetRenderState Lib "dx9_vb" (ByVal pDev As Long, ByVal State As D3DRENDERSTATETYPE, ByVal Value As Long)
Private Declare Sub dev_GetRenderState Lib "dx9_vb" (ByVal pDev As Long, ByVal State As D3DRENDERSTATETYPE, ByRef Value As Long)
Private Declare Sub dev_GetTransform Lib "dx9_vb" (ByVal pDev As Long, ByVal State As D3DTRANSFORMSTATETYPE, ByRef m As D3DMATRIX)
Private Declare Sub dev_SetTransform Lib "dx9_vb" (ByVal pDev As Long, ByVal State As D3DTRANSFORMSTATETYPE, ByRef m As D3DMATRIX)
Private Declare Sub dev_MultiplyTransform Lib "dx9_vb" (ByVal pDev As Long, ByVal State As D3DTRANSFORMSTATETYPE, ByRef m As D3DMATRIX)
Private Declare Sub dev_SetSamplerState Lib "dx9_vb" (ByVal pDev As Long, ByVal Sampler As Long, ByVal sType As D3DSAMPLERSTATETYPE, ByVal Value As Long)
Private Declare Sub dev_GetSamplerState Lib "dx9_vb" (ByVal pDev As Long, ByVal Sampler As Long, ByVal sType As D3DSAMPLERSTATETYPE, ByRef Value As Long)
Private Declare Sub dev_SetTextureStageState Lib "dx9_vb" (ByVal pDev As Long, ByVal Stage As Long, ByVal sType As D3DTEXTURESTAGESTATETYPE, ByVal Value As Long)
Private Declare Sub dev_GetTextureStageState Lib "dx9_vb" (ByVal pDev As Long, ByVal Stage As Long, ByVal sType As D3DTEXTURESTAGESTATETYPE, ByRef Value As Long)
Private Declare Sub dev_CreateVertexBuffer Lib "dx9_vb" (ByVal pDev As Long, ByVal Lenght As Long, ByVal Usage As D3DUSAGE, ByVal FVF As CONST_D3DFVF, ByVal Pool As D3DPOOL, ByRef pVB As Long)
Private Declare Sub dev_SetTexture Lib "dx9_vb" (ByVal pDev As Long, ByVal Stage As Long, ByVal pTex As Long)
Private Declare Sub dev_SetLight Lib "dx9_vb" (ByVal pDev As Long, ByVal index As Long, ByRef pLight As D3DLIGHT9)
Private Declare Sub dev_SetMaterial Lib "dx9_vb" (ByVal pDev As Long, ByRef pMtrl As D3DMATERIAL9)
Private Declare Sub dev_LightEnable Lib "dx9_vb" (ByVal pDev As Long, ByVal LightIndex As Long, ByVal bEnable As Long)
Private Declare Sub dev_CreateIndexBuffer Lib "dx9_vb" (ByVal pDev As Long, ByVal Lenght As Long, ByVal Usage As D3DUSAGE, ByVal iFormat As D3DFORMAT, ByVal Pool As D3DPOOL, ByRef pIB As Long)
Private Declare Sub dev_SetIndices Lib "dx9_vb" (ByVal pDev As Long, ByVal pIB As Long)
Private Declare Sub dev_DrawPrimitive Lib "dx9_vb" (ByVal pDev As Long, ByVal PrimitiveType As D3DPRIMITIVETYPE, ByVal StartVertex As Long, ByVal PrimitiveCount As Long)
Private Declare Sub dev_DrawPrimitiveUp Lib "dx9_vb" (ByVal pDev As Long, ByVal PrimitiveType As D3DPRIMITIVETYPE, ByVal PrimitiveCount As Long, ByVal pVertexStreamZeroData As Long, ByVal VertexStreamZeroStride As Long)
Private Declare Sub dev_DrawIndexedPrimitive Lib "dx9_vb" (ByVal pDev As Long, ByVal PrimitiveType As D3DPRIMITIVETYPE, ByVal BaseVertexIndex As Long, ByVal MinVertexIndex As Long, ByVal NumVertices As Long, ByVal StartIndex As Long, ByVal PrimitiveCount As Long)
Private Declare Sub dev_DrawIndexedPrimitiveUp Lib "dx9_vb" (ByVal pDev As Long, ByVal PrimitiveType As D3DPRIMITIVETYPE, ByVal MinVertexIndex As Long, ByVal NumVertices As Long, ByVal PrimitiveCount As Long, ByVal pIndexData As Long, ByVal IndexDataFormat As Long, ByVal pVertexStreamZeroData As Long, ByVal VertexStreamZeroStride As Long)
Private Declare Sub dev_CreateTexture Lib "dx9_vb" (ByVal pDev As Long, ByVal Width As Long, ByVal Height As Long, ByVal Levels As Long, ByVal Usage As D3DUSAGE, ByVal tFormat As D3DFORMAT, ByVal Pool As D3DPOOL, ByRef pTex As Long)
Private Declare Sub dev_CreateDepthStencilSurface Lib "dx9_vb" (ByVal pDev As Long, ByVal Width As Long, ByVal Height As Long, ByVal tFormat As D3DFORMAT, ByVal MultiSample As D3DMULTISAMPLE_TYPE, ByVal MultiSampleQuality As Long, ByVal Discard As Long, ByRef pSurf As Long)
Private Declare Sub dev_SetRenderTarget Lib "dx9_vb" (ByVal pDev As Long, ByVal RenderTargetIndex As Long, ByVal pSurf As Long)
Private Declare Sub dev_GetRenderTarget Lib "dx9_vb" (ByVal pDev As Long, ByVal RenderTargetIndex As Long, ByRef pSurf As Long)
Private Declare Sub dev_SetDepthStencilSurface Lib "dx9_vb" (ByVal pDev As Long, ByVal pSurf As Long)
Private Declare Sub dev_GetDepthStencilSurface Lib "dx9_vb" (ByVal pDev As Long, ByRef pSurf As Long)
Private Declare Sub dev_CreatePixelShader Lib "dx9_vb" (ByVal pDev As Long, ByVal pFunction As Long, ByRef pShader As Long)
Private Declare Sub dev_SetPixelShader Lib "dx9_vb" (ByVal pDev As Long, ByVal pShader As Long)
Private Declare Sub dev_CreateVertexShader Lib "dx9_vb" (ByVal pDev As Long, ByVal pFunction As Long, ByRef pShader As Long)
Private Declare Sub dev_SetVertexShader Lib "dx9_vb" (ByVal pDev As Long, ByVal pShader As Long)
Private Declare Sub dev_CreateVertexDeclaration Lib "dx9_vb" (ByVal pDev As Long, ByVal pVertexElement As Long, ByRef pDecl As Long)
Private Declare Sub dev_SetVertexDeclaration Lib "dx9_vb" (ByVal pDev As Long, ByVal pDecl As Long)
Private Declare Sub dev_SetPixelShaderConstantI Lib "dx9_vb" (ByVal pDev As Long, ByVal StartRegister As Long, ByVal pData As Long, ByVal v4Count As Long)
Private Declare Sub dev_SetPixelShaderConstantF Lib "dx9_vb" (ByVal pDev As Long, ByVal StartRegister As Long, ByVal pData As Long, ByVal v4fCount As Long)
Private Declare Sub dev_SetPixelShaderConstantB Lib "dx9_vb" (ByVal pDev As Long, ByVal StartRegister As Long, ByVal pData As Long, ByVal BoolCount As Long)
Private Declare Sub dev_SetVertexShaderConstantI Lib "dx9_vb" (ByVal pDev As Long, ByVal StartRegister As Long, ByVal pData As Long, ByVal v4Count As Long)
Private Declare Sub dev_SetVertexShaderConstantF Lib "dx9_vb" (ByVal pDev As Long, ByVal StartRegister As Long, ByVal pData As Long, ByVal v4fCount As Long)
Private Declare Sub dev_SetVertexShaderConstantB Lib "dx9_vb" (ByVal pDev As Long, ByVal StartRegister As Long, ByVal pData As Long, ByVal BoolCount As Long)
Private Declare Sub dev_CreateVolumeTexture Lib "dx9_vb" (ByVal pDev As Long, ByVal Width As Long, ByVal Height As Long, ByVal Depth As Long, ByVal Levels As Long, ByVal Usage As D3DUSAGE, ByVal tFormat As D3DFORMAT, ByVal Pool As D3DPOOL, ByRef pTex As Long)
Private Declare Sub dev_SetClipPlane Lib "dx9_vb" (ByVal pDev As Long, ByVal index As Long, ByVal pPlane As Long)
Private Declare Sub dev_Reset Lib "dx9_vb" (ByVal pDev As Long, ByRef PP As D3DPRESENT_PARAMETERS)
Private Declare Function dev_TestCooperativeLevel Lib "dx9_vb" (ByVal pDev As Long) As HRESULT
Private Declare Function dev_GetRenderTargetData Lib "dx9_vb" (ByVal pDev As Long, ByVal pRenderTarget As Long, ByVal pDestSurface As Long) As HRESULT
Private Declare Sub dev_CreateOffscreenPlainSurface Lib "dx9_vb" (ByVal pDev As Long, ByVal Width As Long, ByVal Height As Long, ByVal tFormat As D3DFORMAT, ByVal Pool As D3DPOOL, ByRef pSurf As Long)
Private mePtr As Long
Friend Function CreateOffscreenPlainSurface(ByVal Width As Long, ByVal Height As Long, ByVal tFormat As D3DFORMAT, ByVal Pool As D3DPOOL) As Direct3DSurface9
   Dim pSurf As Long
   dev_CreateOffscreenPlainSurface Ptr, Width, Height, tFormat, Pool, pSurf
   If pSurf <> 0 Then
      Set CreateOffscreenPlainSurface = New Direct3DSurface9
      CreateOffscreenPlainSurface.Ptr = pSurf
   End If
End Function
Friend Function GetRenderTargetData(RTSurf As Direct3DSurface9, DestSurf As Direct3DSurface9) As HRESULT
   GetRenderTargetData = dev_GetRenderTargetData(Ptr, RTSurf.Ptr, DestSurf.Ptr)
End Function
Friend Function TestCooperativeLevel() As HRESULT
   TestCooperativeLevel = dev_TestCooperativeLevel(Ptr)
End Function
Friend Sub Reset(ByRef PP As D3DPRESENT_PARAMETERS)
   dev_Reset Ptr, PP
End Sub
Friend Sub SetClipPlane(ByVal index As Long, ByVal pPlane As Long)
   dev_SetClipPlane Ptr, index, pPlane
End Sub
Friend Function CreateVolumeTexture(ByVal Width As Long, ByVal Height As Long, ByVal Depth As Long, ByVal Levels As Long, ByVal Usage As D3DUSAGE, ByVal tFormat As D3DFORMAT, ByVal Pool As D3DPOOL) As Direct3DVolumeTexture9
   Dim pTex As Long
   dev_CreateVolumeTexture Ptr, Width, Height, Depth, Levels, Usage, tFormat, Pool, pTex
   If pTex <> 0 Then
      Set CreateVolumeTexture = New Direct3DVolumeTexture9
      CreateVolumeTexture.Ptr = pTex
   End If
End Function
Friend Sub SetPixelShaderConstantI(ByVal StartRegister As Long, ByVal pData As Long, ByVal v4Count As Long)
   dev_SetPixelShaderConstantI Ptr, StartRegister, pData, v4Count
End Sub
Friend Sub SetPixelShaderConstantF(ByVal StartRegister As Long, ByVal pData As Long, ByVal v4fCount As Long)
   dev_SetPixelShaderConstantF Ptr, StartRegister, pData, v4fCount
End Sub
Friend Sub SetPixelShaderConstantB(ByVal StartRegister As Long, ByVal pData As Long, ByVal BoolCount As Long)
   dev_SetPixelShaderConstantB Ptr, StartRegister, pData, BoolCount
End Sub
Friend Sub SetVertexShaderConstantI(ByVal StartRegister As Long, ByVal pData As Long, ByVal v4Count As Long)
   dev_SetVertexShaderConstantI Ptr, StartRegister, pData, v4Count
End Sub
Friend Sub SetVertexShaderConstantF(ByVal StartRegister As Long, ByVal pData As Long, ByVal v4fCount As Long)
   dev_SetVertexShaderConstantF Ptr, StartRegister, pData, v4fCount
End Sub
Friend Sub SetVertexShaderConstantB(ByVal StartRegister As Long, ByVal pData As Long, ByVal BoolCount As Long)
   dev_SetVertexShaderConstantB Ptr, StartRegister, pData, BoolCount
End Sub
Friend Sub SetVertexDeclaration(Decl As Direct3DVertexDeclaration)
   Dim p As Long
   If Decl Is Nothing Then p = 0 Else p = Decl.Ptr
   dev_SetVertexDeclaration Ptr, p
End Sub
Friend Function CreateVertexDeclaration(ByVal pVertexElement As Long) As Direct3DVertexDeclaration
   Dim pDecl As Long
   dev_CreateVertexDeclaration Ptr, pVertexElement, pDecl
   If pDecl <> 0 Then
      Set CreateVertexDeclaration = New Direct3DVertexDeclaration
      CreateVertexDeclaration.Ptr = pDecl
   Else
      Stop
   End If
End Function
Friend Sub SetVertexShader(Shader As Direct3DVertexShader9)
   Dim p As Long
   If Shader Is Nothing Then p = 0 Else p = Shader.Ptr
   dev_SetVertexShader Ptr, p
End Sub
Friend Function CreateVertexShaderFromFile(fName As String) As Direct3DVertexShader9
   Dim pShader As Long
   Dim nf      As Integer
   Dim ar()    As Byte
   nf = FreeFile
   Open fName For Binary As #nf
   ReDim ar(LOF(nf) - 1)
   Get #nf, , ar()
   Close #nf
   dev_CreateVertexShader Ptr, VarPtr(ar(0)), pShader
   If pShader <> 0 Then
      Set CreateVertexShaderFromFile = New Direct3DVertexShader9
      CreateVertexShaderFromFile.Ptr = pShader
   Else
      Stop
   End If
End Function
Friend Sub SetPixelShader(Shader As Direct3DPixelShader9)
   Dim p As Long
   If Shader Is Nothing Then p = 0 Else p = Shader.Ptr
   dev_SetPixelShader Ptr, p
End Sub
Friend Function CreatePixelShaderFromFile(fName As String) As Direct3DPixelShader9
   Dim pShader As Long
   Dim nf      As Integer
   Dim ar()    As Byte
   nf = FreeFile
   Open fName For Binary As #nf
   ReDim ar(LOF(nf) - 1)
   Get #nf, , ar()
   Close #nf
   dev_CreatePixelShader Ptr, VarPtr(ar(0)), pShader
   If pShader <> 0 Then
      Set CreatePixelShaderFromFile = New Direct3DPixelShader9
      CreatePixelShaderFromFile.Ptr = pShader
   Else
      Stop
   End If
End Function
Friend Sub SetDepthStencilSurface(Surf As Direct3DSurface9)
   dev_SetDepthStencilSurface Ptr, Surf.Ptr
End Sub
Friend Function GetDepthStencilSurface() As Direct3DSurface9
   Dim pSurf As Long
   dev_GetDepthStencilSurface Ptr, pSurf
   If pSurf = 0 Then Exit Function
   Set GetDepthStencilSurface = New Direct3DSurface9
   GetDepthStencilSurface.Ptr = pSurf
End Function
Friend Sub SetRenderTarget(ByVal index As Long, Surf As Direct3DSurface9)
   dev_SetRenderTarget Ptr, index, Surf.Ptr
End Sub
Friend Function GetRenderTarget(ByVal index As Long) As Direct3DSurface9
   Dim pSurf As Long
   dev_GetRenderTarget Ptr, index, pSurf
   If pSurf = 0 Then Exit Function
   Set GetRenderTarget = New Direct3DSurface9
   GetRenderTarget.Ptr = pSurf
End Function
Friend Function CreateDepthStencilSurface(ByVal Width As Long, ByVal Height As Long, ByVal tFormat As D3DFORMAT, ByVal MultiSample As D3DMULTISAMPLE_TYPE, ByVal MultiSampleQuality As Long, ByVal Discard As Long) As Direct3DSurface9
   Dim pSurf As Long
   dev_CreateDepthStencilSurface Ptr, Width, Height, tFormat, MultiSample, MultiSampleQuality, Discard, pSurf
   If pSurf <> 0 Then
      Set CreateDepthStencilSurface = New Direct3DSurface9
      CreateDepthStencilSurface.Ptr = pSurf
   End If
End Function
Friend Function CreateTexture(ByVal Width As Long, ByVal Height As Long, ByVal Levels As Long, ByVal Usage As D3DUSAGE, ByVal tFormat As D3DFORMAT, ByVal Pool As D3DPOOL) As Direct3DTexture9
   Dim pTex As Long
   dev_CreateTexture Ptr, Width, Height, Levels, Usage, tFormat, Pool, pTex
   If pTex <> 0 Then
      Set CreateTexture = New Direct3DTexture9
      CreateTexture.Ptr = pTex
   End If
End Function
Friend Sub DrawPrimitive(ByVal PrimitiveType As D3DPRIMITIVETYPE, ByVal StartVertex As Long, ByVal PrimitiveCount As Long)
   dev_DrawPrimitive Ptr, PrimitiveType, StartVertex, PrimitiveCount
End Sub
Friend Sub DrawPrimitiveUp(ByVal PrimitiveType As D3DPRIMITIVETYPE, ByVal PrimitiveCount As Long, ByVal pVertexStreamZeroData As Long, ByVal VertexStreamZeroStride As Long)
   dev_DrawPrimitiveUp Ptr, PrimitiveType, PrimitiveCount, pVertexStreamZeroData, VertexStreamZeroStride
End Sub
Friend Sub DrawIndexedPrimitive(ByVal PrimitiveType As D3DPRIMITIVETYPE, ByVal BaseVertexIndex As Long, ByVal MinVertexIndex As Long, ByVal NumVertices As Long, ByVal StartIndex As Long, ByVal PrimitiveCount As Long)
   dev_DrawIndexedPrimitive Ptr, PrimitiveType, BaseVertexIndex, MinVertexIndex, NumVertices, StartIndex, PrimitiveCount
End Sub
Friend Sub DrawIndexedPrimitiveUp(ByVal PrimitiveType As D3DPRIMITIVETYPE, ByVal MinVertexIndex As Long, ByVal NumVertices As Long, ByVal PrimitiveCount As Long, ByVal pIndexData As Long, ByVal IndexDataFormat As D3DFORMAT, ByVal pVertexStreamZeroData As Long, ByVal VertexStreamZeroStride As Long)
   dev_DrawIndexedPrimitiveUp Ptr, PrimitiveType, MinVertexIndex, NumVertices, PrimitiveCount, pIndexData, IndexDataFormat, pVertexStreamZeroData, VertexStreamZeroStride
End Sub
Friend Sub SetIndices(IB As Direct3DIndexBuffer9)
   dev_SetIndices Ptr, IB.Ptr
End Sub
Friend Function CreateIndexBuffer(ByVal Lenght As Long, ByVal Usage As D3DUSAGE, ByVal iFormat As D3DFORMAT, ByVal Pool As D3DPOOL) As Direct3DIndexBuffer9
   Dim pIB As Long
   dev_CreateIndexBuffer Ptr, Lenght, Usage, iFormat, Pool, pIB
   If pIB <> 0 Then
      Set CreateIndexBuffer = New Direct3DIndexBuffer9
      CreateIndexBuffer.Ptr = pIB
   End If
End Function
Friend Sub LightEnable(ByVal LightIndex As Long, ByVal bEnable As Long)
   dev_LightEnable Ptr, LightIndex, bEnable
End Sub
Friend Sub SetMaterial(ByRef pMtrl As D3DMATERIAL9)
   dev_SetMaterial Ptr, pMtrl
End Sub
Friend Sub SetLight(ByVal index As Long, ByRef pLight As D3DLIGHT9)
   dev_SetLight Ptr, index, pLight
End Sub
Friend Sub SetTextureStageState(ByVal Stage As Long, ByVal sType As D3DTEXTURESTAGESTATETYPE, ByVal Value As Long)
   dev_SetTextureStageState Ptr, Stage, sType, Value
End Sub
Friend Sub GetTextureStageState(ByVal Stage As Long, ByVal sType As D3DTEXTURESTAGESTATETYPE, ByRef Value As Long)
   dev_GetTextureStageState Ptr, Stage, sType, Value
End Sub
Friend Sub SetSamplerState(ByVal Sampler As Long, ByVal sType As D3DSAMPLERSTATETYPE, ByVal Value As Long)
   dev_SetSamplerState Ptr, Sampler, sType, Value
End Sub
Friend Sub GetSamplerState(ByVal Sampler As Long, ByVal sType As D3DSAMPLERSTATETYPE, ByRef Value As Long)
   dev_GetSamplerState Ptr, Sampler, sType, Value
End Sub
Friend Sub SetTransform(ByVal State As D3DTRANSFORMSTATETYPE, ByRef m As D3DMATRIX)
   dev_SetTransform Ptr, State, m
End Sub
Friend Sub GetTransform(ByVal State As D3DTRANSFORMSTATETYPE, ByRef m As D3DMATRIX)
   dev_GetTransform Ptr, State, m
End Sub
Friend Sub MultiplyTransform(ByVal State As D3DTRANSFORMSTATETYPE, ByRef m As D3DMATRIX)
   dev_MultiplyTransform Ptr, State, m
End Sub
Friend Sub SetTexture(ByVal Stage As Long, tex As Direct3DTexture9)
   If tex Is Nothing Then
      dev_SetTexture Ptr, Stage, 0
   Else
      dev_SetTexture Ptr, Stage, tex.Ptr
   End If
End Sub
Friend Sub SetVolumeTexture(ByVal Stage As D3DRENDERSTATETYPE, tex As Direct3DVolumeTexture9)
   If tex Is Nothing Then
      dev_SetTexture Ptr, Stage, 0
   Else
      dev_SetTexture Ptr, Stage, tex.Ptr
   End If
End Sub
Friend Sub SetRenderState(ByVal State As D3DRENDERSTATETYPE, ByVal Value As Long)
   dev_SetRenderState Ptr, State, Value
End Sub
Friend Sub GetRenderState(ByVal State As D3DRENDERSTATETYPE, ByRef Value As Long)
   dev_GetRenderState Ptr, State, Value
End Sub
Friend Sub SetStreamSource(ByVal StreamNumber As Long, VB As Direct3DVertexBuffer9, ByVal Offset As Long, ByVal Stride As Long)
   dev_SetStreamSource Ptr, StreamNumber, VB.Ptr, Offset, Stride
End Sub
Friend Sub SetFVF(ByVal FVF As CONST_D3DFVF)
   dev_SetFVF Ptr, FVF
End Sub
Friend Function CreateVertexBuffer(ByVal Lenght As Long, ByVal Usage As D3DUSAGE, ByVal FVF As CONST_D3DFVF, ByVal Pool As D3DPOOL) As Direct3DVertexBuffer9
   Dim pVB As Long
   dev_CreateVertexBuffer Ptr, Lenght, Usage, FVF, Pool, pVB
   If pVB <> 0 Then
      Set CreateVertexBuffer = New Direct3DVertexBuffer9
      CreateVertexBuffer.Ptr = pVB
   End If
End Function
Friend Sub Clear(ByVal flags As CONST_D3DCLEARFLAGS, ByVal Color As Long, ByVal z As Single, ByVal Stencil As Long)
   dev_Clear Ptr, flags, Color, z, Stencil
End Sub
Friend Function BeginScene() As Boolean
   If dev_BeginScene(Ptr) Then BeginScene = True
End Function
Friend Sub EndScene()
   dev_EndScene Ptr
End Sub
Friend Sub Present()
   dev_Present Ptr
End Sub
Private Sub Class_Terminate()
   dev_Release Ptr
End Sub
Public Property Let Ptr(p As Long)
   If mePtr <> 0 Then dev_Release mePtr
   mePtr = p
End Property
Public Property Get Ptr() As Long
   Ptr = mePtr
End Property
