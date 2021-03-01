VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSky"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type vFormat
    Pos As D3DVECTOR
End Type
Const TriCount As Long = 64
Dim ps         As Direct3DPixelShader9
Dim VS         As Direct3DVertexShader9
Dim vDecl      As Direct3DVertexDeclaration
Dim vBuf       As Direct3DVertexBuffer9
Dim vSize      As Long

Friend Sub BillboardDraw(tex As Direct3DTexture9, Pos As D3DVECTOR, ByVal SizeX As Single, ByVal SizeY As Single, Optional ByVal VFix As Boolean = False)
    Dim Mtrx As D3DMATRIX
    Dim m    As D3DMATRIX
    If VFix Then
        MatrixLookAtRH Mtrx, Pos, Vec3(PlPos.x, Pos.y, PlPos.z), Vec3(0, 1, 0)
    Else
        MatrixLookAtRH Mtrx, Pos, PlPos, Vec3(0, 1, 0)
    End If
    MatrixInverse Mtrx, 0, Mtrx
    MatrixScaling m, SizeX, SizeY, 1
    MatrixMultiply Mtrx, m, Mtrx
    MatrixMultiply Mtrx, Mtrx, mView
    MatrixMultiply Mtrx, Mtrx, mProj
    MatrixTranspose Mtrx, Mtrx
    Dev.SetVertexShaderConstantF 0, VarPtr(Mtrx), 4
    Dev.SetTexture 0, tex
    Dev.SetVertexDeclaration vDecl
    Dev.SetVertexShader VS
    Dev.SetPixelShader ps
    Dev.DrawPrimitiveUp D3DPT_TRIANGLESTRIP, 2, VarPtr(vBuf(0)), vSize
End Sub

Friend Sub Draw(cPos As D3DVECTOR, TexUp As Direct3DTexture9, TexDown As Direct3DTexture9, TexSun As Direct3DTexture9)
    Dim Mtrx As D3DMATRIX
    Dim M1   As D3DMATRIX
    Dim v    As D3DVECTOR
    MatrixLookAtLH Mtrx, Vec3(0, 0, 0), PlDir, PlUp
    MatrixMultiply Mtrx, Mtrx, mProj
    MatrixTranspose VSConst.WVP, Mtrx
    Dev.SetVertexShaderConstantF 0, VarPtr(VSConst), 4
    Dev.SetRenderState D3DRS_ZENABLE, D3DZB_FALSE
    Dev.SetRenderState D3DRS_ALPHABLENDENABLE, D3D_FALSE
    Dev.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    Dev.SetVertexDeclaration vDecl
    Dev.SetStreamSource 0, vBuf, 0, vSize
    Dev.SetVertexShader VS
    Dev.SetPixelShader ps
    Dev.SetTexture 0, TexUp
    Dev.DrawPrimitive D3DPT_TRIANGLEFAN, 0, TriCount
    MatrixScaling M1, 1, -0.5, 1
    MatrixMultiply Mtrx, M1, Mtrx
    MatrixTranspose VSConst.WVP, Mtrx
    Dev.SetVertexShaderConstantF 0, VarPtr(VSConst), 4
    Dev.SetRenderState D3DRS_CULLMODE, D3DCULL_CW
    Dev.SetTexture 0, TexDown
    Dev.DrawPrimitive D3DPT_TRIANGLEFAN, 0, TriCount
    Dev.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    Dev.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
    Dev.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    Dev.SetRenderState D3DRS_ALPHABLENDENABLE, D3D_TRUE
    Vec3Add v, PlPos, SunLightPos
    '    BillboardDraw TexSun, v, 1, 1
End Sub

Private Sub Class_Initialize()
    Dim n      As Integer
    Dim tmp    As Single
    Dim Vert() As vFormat
    Dim m      As D3DMATRIX
    tmp = 2 * Pi / TriCount
    ReDim Vert(TriCount + 1)
    vSize = Len(Vert(0))
    Set vBuf = Dev.CreateVertexBuffer((TriCount + 2) * vSize, D3DUSAGE_NONE, D3DFVF_XYZ Or D3DFVF_DIFFUSE, D3DPOOL_MANAGED)
    Vert(0).Pos = Vec3(0, 1, 0)
    For n = 1 To TriCount + 1
        Vert(n).Pos = Vec3(Sin(n * tmp), -0.001, -Cos(n * tmp))
    Next n
    vBuf.SetData 0, (TriCount + 2) * vSize, VarPtr(Vert(0)), 0
    VS_Create
    PS_Create
End Sub

Private Sub VS_Create()
    Dim vsDecl(1) As D3DVERTEXELEMENT9
    vsDecl(0) = VertexElement(0, 0, D3DDECLTYPE_FLOAT3, D3DDECLMETHOD_DEFAULT, D3DDECLUSAGE_POSITION, 0)
    vsDecl(1) = VertexElementEnd
    Set vDecl = Dev.CreateVertexDeclaration(VarPtr(vsDecl(0)))
    Set VS = Dev.CreateVertexShaderFromFile(App.Path & "\3DEngine\ShaderCompiler\Out\Sky.vsh.shader")
End Sub

Private Sub PS_Create()
    Set ps = Dev.CreatePixelShaderFromFile(App.Path & "\3DEngine\ShaderCompiler\Out\Sky.psh.shader")
End Sub

Private Sub Class_Terminate()
    Set VS = Nothing
    Set ps = Nothing
    Set vDecl = Nothing
    Set vBuf = Nothing
End Sub
