VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLandScape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type vFormat
   Pos As D3DVECTOR
   Normal As D3DVECTOR
End Type
Dim vBuf     As Direct3DVertexBuffer9
Dim iBuf     As Direct3DIndexBuffer9
Dim ps       As Direct3DPixelShader9
Dim VS       As Direct3DVertexShader9
Dim vDecl    As Direct3DVertexDeclaration
Dim Vert()   As vFormat
Dim vSize    As Long
Dim vCnt     As Long
Dim iCnt     As Long
Dim tCnt     As Long
Dim xMax     As Single, xMin As Single
Dim yMax     As Single, yMin As Single
Dim zMax     As Single, zMin As Single
Dim meHScale As Single, meSize As Long, InvHScale As Single
Private Const FVF = D3DFVF_XYZ
Public Sub Draw()
   Dim Mtrx As D3DMATRIX
   VSConst.LitePos = SunLightPos
   MatrixMultiply Mtrx, mView, mProj
   MatrixTranspose VSConst.WVP, Mtrx
   MatrixMultiply Mtrx, mViewShad, mProjShad
   MatrixTranspose VSConst.mShad, Mtrx
   Dev.SetVertexShaderConstantF 0, VarPtr(VSConst), 16
   Dev.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
   Dev.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE
   Dev.SetRenderState D3DRS_ZWRITEENABLE, D3D_TRUE
   Dev.SetRenderState D3DRS_ALPHABLENDENABLE, D3D_FALSE
   Dev.SetStreamSource 0, vBuf, 0, vSize
   Dev.SetIndices iBuf
   Dev.SetVertexDeclaration vDecl
   Dev.SetVertexShader VS
   Dev.SetPixelShader ps
   Dev.DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, 0, vCnt, 0, tCnt
End Sub
Public Sub DrawShade()
   Dim Mtrx As D3DMATRIX
   MatrixMultiply Mtrx, mViewShad, mProjShad
   MatrixTranspose VSConst.WVP, Mtrx
   Dev.SetVertexShaderConstantF 0, VarPtr(VSConst), 4
   Dev.SetVertexDeclaration vDecl
   Dev.SetStreamSource 0, vBuf, 0, vSize
   Dev.SetIndices iBuf
   Dev.DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, 0, vCnt, 0, tCnt
End Sub
Public Sub Init(fName As String, ByVal HScale As Single, ByVal vScale As Single)
   Dim Ind() As Integer
   Dim b()   As Byte
   Dim x     As Long, z As Long
   Dim X1    As Long, z1 As Long
   Dim X2    As Long, z2 As Long
   Dim v1    As D3DVECTOR
   Dim v2    As D3DVECTOR
   Dim nf    As Integer
   vCnt = 129 * 129
   tCnt = 128 * 128 * 2&
   iCnt = tCnt * 3
   meSize = 128
   xMin = 0
   yMin = 0
   zMin = 0
   xMax = meSize * HScale
   yMax = 255 * vScale
   zMax = meSize * HScale
   meHScale = HScale
   InvHScale = 1 / HScale
   ReDim b(meSize - 1, meSize - 1)
   ReDim Vert(meSize, meSize)
   ReDim Ind(iCnt - 1)
   nf = FreeFile
   Open fName For Binary As #nf
   Get #nf, 19, b()
   Close #nf
   For z = 0 To 128
      For x = 0 To 128
         X1 = x - 1: If X1 < 0 Then X1 = 0 Else If X1 > 127 Then X1 = 127
         X2 = x: If X2 > 127 Then X2 = 127
         z1 = z - 1: If z1 < 0 Then z1 = 0 Else If z1 > 127 Then z1 = 127
         z2 = z: If z2 > 127 Then z2 = 127
         Vert(x, z).Pos = Vec3(x * HScale, (0& + b(X1, z1) + b(X1, z2) + b(X2, z1) + b(X2, z2)) * vScale / (255 * 4), z * HScale)
      Next x
   Next z
   For z = 0 To 128
      For x = 0 To 128
         If x < 1 Then X1 = 0 Else X1 = x - 1
         If x > 127 Then X2 = 128 Else X2 = x + 1
         If z < 1 Then z1 = 0 Else z1 = z - 1
         If z > 127 Then z2 = 128 Else z2 = z + 1
         Vec3Subtract v1, Vert(X2, z).Pos, Vert(X1, z).Pos
         Vec3Subtract v2, Vert(x, z2).Pos, Vert(x, z1).Pos
         Vec3Cross Vert(x, z).Normal, v2, v1
         Vec3Normalize Vert(x, z).Normal, Vert(x, z).Normal
      Next x
   Next z
   For z = 0 To 127
      For x = 0 To 127
         Ind((x + z * 128) * 6 + 0) = x + z * 129 + 0
         Ind((x + z * 128) * 6 + 1) = x + z * 129 + 129
         Ind((x + z * 128) * 6 + 2) = x + z * 129 + 130
         Ind((x + z * 128) * 6 + 3) = x + z * 129 + 0
         Ind((x + z * 128) * 6 + 4) = x + z * 129 + 130
         Ind((x + z * 128) * 6 + 5) = x + z * 129 + 1
      Next x
   Next z
   Set vBuf = Nothing
   vSize = Len(Vert(0, 0))
   Set vBuf = Dev.CreateVertexBuffer(vCnt * vSize, D3DUSAGE_NONE, FVF, D3DPOOL_MANAGED)
   vBuf.SetData 0, vCnt * vSize, VarPtr(Vert(0, 0)), 0
   Set iBuf = Nothing
   Set iBuf = Dev.CreateIndexBuffer(iCnt * Len(Ind(0)), D3DUSAGE_NONE, D3DFMT_INDEX16, D3DPOOL_MANAGED)
   iBuf.SetData 0, iCnt * Len(Ind(0)), VarPtr(Ind(0)), 0
   VS_Create
   PS_Create
End Sub
Friend Sub GetHeight(ByVal x As Single, ByVal z As Single, RetHeight As Single)
   Dim xx As Single, zz As Single, iX As Long, iZ As Long
   xx = (x - xMin) * InvHScale
   zz = (z - zMin) * InvHScale
   iX = Int(xx)
   iZ = Int(zz)
   xx = xx - iX
   zz = zz - iZ
   If iX < 0 Or iX >= meSize Or iZ < 0 Or iZ >= meSize Then RetHeight = -999999: Exit Sub
   RetHeight = Vert(iX, iZ).Pos.y * (1 - xx) * (1 - zz) + Vert(iX + 1, iZ).Pos.y * xx * (1 - zz) + Vert(iX, iZ + 1).Pos.y * (1 - xx) * zz + Vert(iX + 1, iZ + 1).Pos.y * xx * zz
End Sub
Friend Sub GetNorm(ByVal x As Single, ByVal z As Single, RetN As D3DVECTOR)
   Dim xx As Single, zz As Single, iX As Long, iZ As Long
   Dim v1 As D3DVECTOR, v2 As D3DVECTOR
   xx = (x - xMin) * InvHScale
   zz = (z - zMin) * InvHScale
   iX = Int(xx)
   iZ = Int(zz)
   xx = xx - iX
   zz = zz - iZ
   If iX < 0 Or iX >= meSize Or iZ < 0 Or iZ >= meSize Then Exit Sub
   Vec3Subtract v1, Vert(iX + 1, iZ).Pos, Vert(iX, iZ + 1).Pos
   Vec3Subtract v2, Vert(iX, iZ).Pos, Vert(iX + 1, iZ + 1).Pos
   Vec3Cross RetN, v1, v2
   Vec3Normalize RetN, RetN
End Sub
Friend Function RayInterSector(RayPos As D3DVECTOR, RayDir As D3DVECTOR, RetV As D3DVECTOR, RetN As D3DVECTOR) As Boolean
   Dim v As D3DVECTOR, v1 As D3DVECTOR, s As Single, sOld As Single
   Vec3Scale v, RayDir, meHScale / Sqr(RayDir.x * RayDir.x + RayDir.z * RayDir.z + 0.000001)
   v1 = RayPos
   GetHeight v1.x, v1.z, s
   sOld = v1.y - s
   Do
      Vec3Add v1, v1, v
      GetHeight v1.x, v1.z, s
      s = v1.y - s
      If s < 0 Then Exit Do
      If v1.x < xMin Then Exit Function
      If v1.x > xMax Then Exit Function
      If v1.z < zMin Then Exit Function
      If v1.z > zMax Then Exit Function
      If v1.y > yMax And v1.y > RayPos.y Then Exit Function
      sOld = s
   Loop
   s = s / (sOld - s)
   Vec3Scale v, v, s
   Vec3Add RetV, v1, v
   GetNorm RetV.x, RetV.z, RetN
   RayInterSector = True
End Function
Private Sub VS_Create()
   Dim vsDecl(2) As D3DVERTEXELEMENT9
   vsDecl(0) = VertexElement(0, 0, D3DDECLTYPE_FLOAT3, D3DDECLMETHOD_DEFAULT, D3DDECLUSAGE_POSITION, 0)
   vsDecl(1) = VertexElement(0, 12, D3DDECLTYPE_FLOAT3, D3DDECLMETHOD_DEFAULT, D3DDECLUSAGE_NORMAL, 0)
   vsDecl(2) = VertexElementEnd
   Set vDecl = Dev.CreateVertexDeclaration(VarPtr(vsDecl(0)))
   Set VS = Dev.CreateVertexShaderFromFile(App.Path & "\3DEngine\ShaderCompiler\Out\Land.vsh.shader")
End Sub
Private Sub PS_Create()
   Set ps = Dev.CreatePixelShaderFromFile(App.Path & "\3DEngine\ShaderCompiler\Out\Land.psh.shader")
End Sub
Private Sub Class_Terminate()
   Set ps = Nothing
   Set VS = Nothing
   Set vDecl = Nothing
   Set vBuf = Nothing
   Set iBuf = Nothing
End Sub
