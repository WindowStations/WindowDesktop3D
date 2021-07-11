VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type vFormat
   Pos As D3DVECTOR2
   tu0 As Single
   tv0 As Single
End Type
Private vBuf(3) As vFormat
Private ps      As Direct3DPixelShader9
Private VS      As Direct3DVertexShader9
Private vDecl   As Direct3DVertexDeclaration
Private kw      As Single
Private kh      As Single
Private vSize   As Long
Friend Sub Draw(tex As Direct3DTexture9, ByVal PosX As Single, ByVal PosY As Single, ByVal Width As Single, ByVal Height As Single, Optional ByVal a As Single = 1)
   On Error Resume Next
   Dim w  As Single
   Dim h  As Single
   Dim v4 As D3DVECTOR4
   w = Width * 0.5
   h = Height * 0.5
   vBuf(0).Pos.x = -1 + (PosX - w) * kw
   vBuf(0).Pos.y = 1 - (PosY + h) * kh
   vBuf(1).Pos.x = -1 + (PosX - w) * kw
   vBuf(1).Pos.y = 1 - (PosY - h) * kh
   vBuf(2).Pos.x = -1 + (PosX + w) * kw
   vBuf(2).Pos.y = 1 - (PosY + h) * kh
   vBuf(3).Pos.x = -1 + (PosX + w) * kw
   vBuf(3).Pos.y = 1 - (PosY - h) * kh
   v4 = Vec4(0, 0, 0, a * 2)
   Dev.SetPixelShaderConstantF 0, VarPtr(v4), 1
   Dev.SetTexture 1, tex
   Dev.SetVertexDeclaration vDecl
   Dev.SetVertexShader VS
   Dev.SetPixelShader ps
   Dev.DrawPrimitiveUp D3DPT_TRIANGLESTRIP, 2, VarPtr(vBuf(0)), vSize
End Sub
Friend Sub Initialize()
   On Error Resume Next
   kw = 2 / BBWidth
   kh = 2 / BBHeight
   VS_Create
   PS_Create
   vSize = Len(vBuf(0))
   vBuf(0).tu0 = 0
   vBuf(0).tv0 = 1
   vBuf(1).tu0 = 0
   vBuf(1).tv0 = 0
   vBuf(2).tu0 = 1
   vBuf(2).tv0 = 1
   vBuf(3).tu0 = 1
   vBuf(3).tv0 = 0
End Sub
Private Sub VS_Create()
   On Error Resume Next
   Dim vsDecl(2) As D3DVERTEXELEMENT9
   vsDecl(0) = VertexElement(0, 0, D3DDECLTYPE_FLOAT2, D3DDECLMETHOD_DEFAULT, D3DDECLUSAGE_POSITION, 0)
   vsDecl(1) = VertexElement(0, 8, D3DDECLTYPE_FLOAT2, D3DDECLMETHOD_DEFAULT, D3DDECLUSAGE_TEXCOORD, 0)
   vsDecl(2) = VertexElementEnd
   Set vDecl = Dev.CreateVertexDeclaration(VarPtr(vsDecl(0)))
   Set VS = Dev.CreateVertexShaderFromFile(App.Path & "\3DEngine\ShaderCompiler\Out\Sprite.vsh.shader")
End Sub
Private Sub PS_Create()
   On Error Resume Next
   Set ps = Dev.CreatePixelShaderFromFile(App.Path & "\3DEngine\ShaderCompiler\Out\Sprite.psh.shader")
End Sub
Friend Sub Terminate()
   Set ps = Nothing
   Set VS = Nothing
   Set vDecl = Nothing
End Sub
