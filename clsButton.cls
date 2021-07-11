VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ButtonMesh As clsMesh
Private vDecl      As Direct3DVertexDeclaration
Private VS         As Direct3DVertexShader9
Private ps         As Direct3DPixelShader9
Public Event OnPush(ByVal iButton As Long)
Public Event OnStruck(ByVal iButton As Long, ByVal jProjectile As Long)
Public Sub Push(ByVal iButton As Long)
   On Error Resume Next
   RaiseEvent OnPush(iButton)
End Sub
Public Sub Struck(ByVal iButton As Long, ByVal jProjectile As Long)
   On Error Resume Next
   RaiseEvent OnStruck(iButton, jProjectile)
End Sub
Friend Sub Save(ByVal nf As Long)
   On Error Resume Next
   Dim i As Long
   Put #nf, , DoorCnt
   For i = 1 To DoorCnt
      Put #nf, , Door(i)
   Next i
End Sub
Friend Sub Load(ByVal nf As Long)
   On Error Resume Next
   Dim i As Long
   Get #nf, , DoorCnt
   For i = 1 To DoorCnt
      Get #nf, , Door(i)
      Sector(Door(i).Sect).yMin = Door(i).yMin + Door(i).PosH
      Sector(Door(i).Sect).yMax = Door(i).yMax + Door(i).PosH
      Door(i).Pos.y = (Sector(Door(i).Sect).yMin + Sector(Door(i).Sect).yMax) * 0.5
      Sector(Door(i).Sect).yMid = Door(i).Pos.y
   Next i
End Sub
Friend Sub Tick()
   On Error Resume Next
   Dim i    As Long
   Dim f    As Boolean
   Dim v    As D3DVECTOR
   Dim m    As D3DMATRIX
   Static t As Long
   For i = 1 To DoorCnt
      If Door(i).Spd <> 0 Then
         Door(i).PosH = Door(i).PosH + Door(i).Spd
         If Door(i).PosH <= Door(i).pMin Then
            Door(i).PosH = Door(i).pMin
            Door(i).Spd = 0
         ElseIf Door(i).PosH >= Door(i).pMax Then
            Door(i).PosH = Door(i).pMax
            Door(i).Spd = 0
         End If
         Sector(Door(i).Sect).yMin = Door(i).yMin + Door(i).PosH
         Sector(Door(i).Sect).yMax = Door(i).yMax + Door(i).PosH
         Door(i).Pos.y = (Sector(Door(i).Sect).yMin + Sector(Door(i).Sect).yMax) * 0.5
         Sector(Door(i).Sect).yMid = Door(i).Pos.y
         MatrixTranslation Door(i).Mtrx, 0, Door(i).PosH, 0
      End If
   Next
   Dim j  As Long
   Dim v2 As D3DVECTOR
   For i = 0 To ButCnt - 1
      f = True
      If But(i).Col.x > 0 Then But(i).Col.x = But(i).Col.x - 0.02: f = False
      If But(i).Col.y > 0 Then But(i).Col.y = But(i).Col.y - 0.02: f = False
      If But(i).Col.z > 0 Then But(i).Col.z = But(i).Col.z - 0.05: f = False
      If f And t = 0 Then But(i).Col.z = 1
      If But(i).Pos.y <> Sector(But(i).Sect).yMid Then
         But(i).Pos.y = Sector(But(i).Sect).yMid
         Vec3Add v, But(i).Pos, But(i).Norm
         MatrixLookAtLH m, But(i).Pos, v, Vec3(0, 1, 0)
         MatrixInverse But(i).Mtrx, 0, m
      End If
      v2 = v
      Vec3Subtract v2, But(i).Pos, Projectiles.Pos 'j
      If Vec3LengthSq(v2) < 3 Then
         Projectiles.Timer = 0 'j
         frmMain.ProjectileClass.Remove j
         Struck i, j
      End If
   Next
   t = (t + 1) And 127
End Sub
Friend Sub Use()
   On Error Resume Next
   Dim i As Long
   Dim d As Long
   Dim v As D3DVECTOR
   Dim s As Single
   Dim f As Boolean
   For i = 0 To ButCnt - 1
      Vec3Subtract v, But(i).Pos, PlPos
      If Vec3LengthSq(v) < 4 Then
         Push i
         s = Vec3Dot(But(i).Norm, PlDir)
         If s < -0.2 Then
            For d = 1 To DoorCnt
               If But(i).Target = Door(d).Tag Then
                  '                        If Door(d).Spd = 0 Then
                  '                            If Door(d).PosH = Door(d).pMin Then
                  '                                Door(d).Spd = 0.01
                  '                            Else
                  '                                Door(d).Spd = -0.01
                  '                            End If
                  'frmMain.soundClass.DoorOperate (Door(d).pMax - Door(d).pMin) * 100, d
                  f = True
                  '                        End If
                  frmMain.TerminateEXE "WindowContextMenu.exe"
                  Unload frmMain
               End If
            Next d
            If f Then
               But(i).Col = Vec4(0.6, 1, 0.4, 0)
               'frmMain.soundClass.ButtonOn
            Else
               But(i).Col = Vec4(0.6, 0, 0, 0)
               'frmMain.soundClass.ButtonFail
            End If
         End If
      End If
   Next
End Sub
Friend Sub add(Vert3() As Map3Vertex, ByVal t As Long, ByVal s As Long)
   Dim v As D3DVECTOR
   Dim m As D3DMATRIX
   ReDim Preserve But(ButCnt)
   But(ButCnt).Norm = Vert3(0).Norm
   Vec3Add v, Vert3(0).Pos, Vert3(1).Pos
   Vec3Add v, v, Vert3(2).Pos
   Vec3Add v, v, Vert3(3).Pos
   Vec3Scale But(ButCnt).Pos, v, 0.25
   Vec3Add v, But(ButCnt).Pos, But(ButCnt).Norm
   MatrixLookAtLH m, But(ButCnt).Pos, v, Vec3(0, 1, 0)
   MatrixInverse But(ButCnt).Mtrx, 0, m
   But(ButCnt).Target = t
   But(ButCnt).Sect = s
   ButCnt = ButCnt + 1
End Sub
Friend Sub Draw()
   Dim Mtrx As D3DMATRIX
   Dim v3   As D3DVECTOR
   Dim i    As Long
   Dev.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE
   Dev.SetRenderState D3DRS_ZWRITEENABLE, D3D_TRUE
   Dev.SetRenderState D3DRS_ALPHABLENDENABLE, D3D_FALSE
   Dev.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
   Dev.SetVertexDeclaration vDecl
   Dev.SetVertexShader VS
   Dev.SetPixelShader ps
   For i = 0 To ButCnt - 1
      MatrixTranspose Mtrx, But(i).Mtrx
      Vec3Subtract v3, PlPos, But(i).Pos
      Vec3TransformNormal VSConst.CamPos, v3, Mtrx
      Vec3TransformNormal VSConst.LitePos, SunLightPos, Mtrx
      MatrixMultiply Mtrx, But(i).Mtrx, mView
      MatrixMultiply Mtrx, Mtrx, mProj
      MatrixTranspose VSConst.WVP, Mtrx
      MatrixTranspose VSConst.World, But(i).Mtrx
      MatrixMultiply Mtrx, But(i).Mtrx, mViewShad
      MatrixMultiply Mtrx, Mtrx, mProjShad
      MatrixTranspose VSConst.mShad, Mtrx
      Dev.SetVertexShaderConstantF 0, VarPtr(VSConst), 16
      Dev.SetPixelShaderConstantF 3, VarPtr(But(i).Col), 1
      ButtonMesh.Draw
   Next i
End Sub
Friend Sub Initialize()
   Set ButtonMesh = New clsMesh
   ButtonMesh.CreateButton
   VSCreate
   PSCreate
End Sub
Private Sub VSCreate()
   Dim vsDecl(3) As D3DVERTEXELEMENT9
   vsDecl(0) = VertexElement(0, 0, D3DDECLTYPE_FLOAT3, D3DDECLMETHOD_DEFAULT, D3DDECLUSAGE_POSITION, 0)
   vsDecl(1) = VertexElement(0, 12, D3DDECLTYPE_FLOAT3, D3DDECLMETHOD_DEFAULT, D3DDECLUSAGE_NORMAL, 0)
   vsDecl(2) = VertexElement(0, 24, D3DDECLTYPE_FLOAT2, D3DDECLMETHOD_DEFAULT, D3DDECLUSAGE_TEXCOORD, 0)
   vsDecl(3) = VertexElementEnd
   Set vDecl = Dev.CreateVertexDeclaration(VarPtr(vsDecl(0)))
   Set VS = Dev.CreateVertexShaderFromFile(App.Path & "\3DEngine\ShaderCompiler\Out\Button.vsh.shader")
End Sub
Private Sub PSCreate()
   Set ps = Dev.CreatePixelShaderFromFile(App.Path & "\3DEngine\ShaderCompiler\Out\Button.psh.shader")
End Sub
Friend Sub Terminate()
   Set ButtonMesh = Nothing
   Set VS = Nothing
   Set ps = Nothing
   Set vDecl = Nothing
End Sub
