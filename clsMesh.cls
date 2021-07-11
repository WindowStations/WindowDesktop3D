VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMesh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type vFormat
   Pos As D3DVECTOR
   Normal As D3DVECTOR
   tu As Single
   tv As Single
End Type
Private vBuf  As Direct3DVertexBuffer9
Private iBuf  As Direct3DIndexBuffer9
Private vSize As Long
Private vCnt  As Long
Private iCnt  As Long
Private tCnt  As Long
Private Const FVF = D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1
Public xMax As Single
Public xMin As Single
Public yMax As Single
Public yMin As Single
Public zMax As Single
Public zMin As Single
Public Sub Draw()
   On Error Resume Next
   Dev.SetStreamSource 0, vBuf, 0, vSize
   Dev.SetIndices iBuf
   Dev.DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, 0, vCnt, 0, tCnt
End Sub
Friend Sub CreateSphere(ByVal xCnt As Long, ByVal yCnt As Long, ByVal Size As Single)
   On Error Resume Next
   Dim Vert() As vFormat
   Dim Ind()  As Integer
   Dim x      As Long
   Dim y      As Long
   Dim SinX   As Single
   Dim CosX   As Single
   Dim SinY   As Single
   Dim CosY   As Single
   Dim nx     As Single
   Dim ny     As Single
   Dim nz     As Single
   vCnt = (yCnt - 1) * xCnt + 2
   tCnt = xCnt * (yCnt - 1) * 2
   iCnt = tCnt * 3
   ReDim Vert(vCnt - 1)
   vSize = Len(Vert(0))
   Vert(0) = Vertex(0, Size, 0, 0, 1, 0)
   For y = 1 To yCnt - 1
      SinY = Sin((yCnt / 2 - y) * Pi / yCnt)
      CosY = Cos((yCnt / 2 - y) * Pi / yCnt)
      For x = 0 To xCnt - 1
         SinX = Sin(x * 2 * Pi / xCnt)
         CosX = Cos(x * 2 * Pi / xCnt)
         nx = SinX * CosY
         ny = SinY
         nz = CosX * CosY
         Vert((y - 1) * xCnt + x + 1) = Vertex(nx * Size, ny * Size, nz * Size, nx, ny, nz)
      Next x
   Next y
   Vert(vCnt - 1) = Vertex(0, -Size, 0, 0, -1, 0)
   ReDim Ind(iCnt - 1)
   For x = 0 To xCnt - 1
      Ind(x * 3 + 0) = 0
      Ind(x * 3 + 1) = x + 1
      Ind(x * 3 + 2) = ((x + 1) Mod xCnt) + 1
   Next x
   For y = 0 To yCnt - 3
      For x = 0 To xCnt - 1
         Ind((x + y * xCnt) * 6 + xCnt * 3 + 0) = xCnt * (y + 0) + x + 1
         Ind((x + y * xCnt) * 6 + xCnt * 3 + 1) = xCnt * (y + 1) + ((x + 1) Mod xCnt) + 1
         Ind((x + y * xCnt) * 6 + xCnt * 3 + 2) = xCnt * (y + 0) + ((x + 1) Mod xCnt) + 1
         Ind((x + y * xCnt) * 6 + xCnt * 3 + 3) = xCnt * (y + 0) + x + 1
         Ind((x + y * xCnt) * 6 + xCnt * 3 + 4) = xCnt * (y + 1) + x + 1
         Ind((x + y * xCnt) * 6 + xCnt * 3 + 5) = xCnt * (y + 1) + ((x + 1) Mod xCnt) + 1
      Next x
   Next y
   For x = 0 To xCnt - 1
      Ind(xCnt * (yCnt - 1) * 6 + (x - xCnt) * 3 + 0) = vCnt - 1
      Ind(xCnt * (yCnt - 1) * 6 + (x - xCnt) * 3 + 1) = vCnt + (((x + 1) Mod xCnt) - xCnt) - 1
      Ind(xCnt * (yCnt - 1) * 6 + (x - xCnt) * 3 + 2) = vCnt + (x - xCnt) - 1
   Next x
   Set vBuf = Nothing
   Set vBuf = Dev.CreateVertexBuffer(vCnt * vSize, D3DUSAGE_NONE, FVF, D3DPOOL_MANAGED)
   vBuf.SetData 0, vCnt * vSize, VarPtr(Vert(0)), 0
   Set iBuf = Nothing
   Set iBuf = Dev.CreateIndexBuffer(iCnt * Len(Ind(0)), D3DUSAGE_NONE, D3DFMT_INDEX16, D3DPOOL_MANAGED)
   iBuf.SetData 0, iCnt * Len(Ind(0)), VarPtr(Ind(0)), 0
End Sub
Private Function Vertex(ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal nx As Single, ByVal ny As Single, ByVal nz As Single) As vFormat
   On Error Resume Next
   Vertex.Pos = Vec3(x, y, z)
   Vertex.Normal = Vec3(nx, ny, nz)
End Function
Friend Sub CreateButton()
   On Error Resume Next
   Dim Vert() As vFormat
   Dim Ind()  As Integer
   Dim v      As D3DVECTOR
   Dim i      As Long
   Const sz   As Single = 0.5
   vCnt = 12
   tCnt = 14
   iCnt = tCnt * 3
   ReDim Vert(vCnt - 1)
   ReDim Ind(iCnt - 1)
   vSize = Len(Vert(0))
   Vert(0).Pos = Vec3(sz * 2, sz, 0)
   Vert(1).Pos = Vec3(sz, sz, 0.1)
   Vert(2).Pos = Vec3(sz, sz * 2, 0)
   Vert(3).Pos = Vec3(-sz, sz * 2, 0)
   Vert(4).Pos = Vec3(-sz, sz, 0.1)
   Vert(5).Pos = Vec3(-sz * 2, sz, 0)
   Vert(6).Pos = Vec3(-sz * 2, -sz, 0)
   Vert(7).Pos = Vec3(-sz, -sz, 0.1)
   Vert(8).Pos = Vec3(-sz, -sz * 2, 0)
   Vert(9).Pos = Vec3(sz, -sz * 2, 0)
   Vert(10).Pos = Vec3(sz, -sz, 0.1)
   Vert(11).Pos = Vec3(sz * 2, -sz, 0)
   For i = 0 To 11
      Vec3Subtract v, Vert(i).Pos, Vec3(0, 0, -5)
      Vec3Normalize Vert(i).Normal, v
      Vert(i).tu = 0.5 - Vert(i).Pos.x * 0.5
      Vert(i).tv = 0.5 - Vert(i).Pos.y * 0.5
   Next i
   For i = 0 To 3
      Ind(i * 9 + 0) = (i * 3 + 0) Mod 12
      Ind(i * 9 + 1) = (i * 3 + 2) Mod 12
      Ind(i * 9 + 2) = (i * 3 + 1) Mod 12
      Ind(i * 9 + 3) = (i * 3 + 1) Mod 12
      Ind(i * 9 + 4) = (i * 3 + 2) Mod 12
      Ind(i * 9 + 5) = (i * 3 + 3) Mod 12
      Ind(i * 9 + 6) = (i * 3 + 1) Mod 12
      Ind(i * 9 + 7) = (i * 3 + 3) Mod 12
      Ind(i * 9 + 8) = (i * 3 + 4) Mod 12
   Next i
   Ind(36) = 1
   Ind(37) = 4
   Ind(38) = 7
   Ind(39) = 1
   Ind(40) = 7
   Ind(41) = 10
   Set vBuf = Nothing
   Set vBuf = Dev.CreateVertexBuffer(vCnt * vSize, D3DUSAGE_NONE, FVF, D3DPOOL_MANAGED)
   vBuf.SetData 0, vCnt * vSize, VarPtr(Vert(0)), 0
   Set iBuf = Nothing
   Set iBuf = Dev.CreateIndexBuffer(iCnt * Len(Ind(0)), D3DUSAGE_NONE, D3DFMT_INDEX16, D3DPOOL_MANAGED)
   iBuf.SetData 0, iCnt * Len(Ind(0)), VarPtr(Ind(0)), 0
End Sub
Friend Sub Init(fName As String, ByVal sc As Single, AddV As D3DVECTOR)
   On Error Resume Next
   Dim Vert() As vFormat
   Dim Ind()  As Integer
   Dim nf     As Integer
   Dim i      As Long
   nf = FreeFile
   Open fName For Binary As #nf
   Get #nf, , vCnt
   Get #nf, , iCnt
   ReDim Vert(vCnt - 1)
   ReDim Ind(iCnt - 1)
   Get #nf, , Vert()
   Get #nf, , Ind()
   Close #nf
   yMin = 1000000
   yMax = -1000000
   xMin = 1000000
   xMax = -1000000
   zMin = 1000000
   zMax = -1000000
   vSize = Len(Vert(0))
   For i = 0 To vCnt - 1
      Vert(i).Pos.x = Vert(i).Pos.x * sc + AddV.x
      Vert(i).Pos.y = Vert(i).Pos.y * sc + AddV.y
      Vert(i).Pos.z = Vert(i).Pos.z * sc + AddV.z
      If xMin > Vert(i).Pos.x Then xMin = Vert(i).Pos.x
      If xMax < Vert(i).Pos.x Then xMax = Vert(i).Pos.x
      If yMin > Vert(i).Pos.y Then yMin = Vert(i).Pos.y
      If yMax < Vert(i).Pos.y Then yMax = Vert(i).Pos.y
      If zMin > Vert(i).Pos.z Then zMin = Vert(i).Pos.z
      If zMax < Vert(i).Pos.z Then zMax = Vert(i).Pos.z
   Next i
   Set vBuf = Nothing
   Set vBuf = Dev.CreateVertexBuffer(vCnt * vSize, D3DUSAGE_NONE, FVF, D3DPOOL_MANAGED)
   vBuf.SetData 0, vCnt * vSize, VarPtr(Vert(0)), 0
   tCnt = iCnt \ 3
   Set iBuf = Nothing
   Set iBuf = Dev.CreateIndexBuffer(iCnt * Len(Ind(0)), D3DUSAGE_NONE, D3DFMT_INDEX16, D3DPOOL_MANAGED)
   iBuf.SetData 0, iCnt * Len(Ind(0)), VarPtr(Ind(0)), 0
End Sub
Private Sub Class_Terminate()
   Set vBuf = Nothing
   Set iBuf = Nothing
End Sub
