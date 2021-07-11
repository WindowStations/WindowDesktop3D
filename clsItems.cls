VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private vSize         As Long
Private Sphere        As clsMesh
Private Mesh()        As clsMesh
Private MeshCnt       As Long
Private ChangePhase   As Double
'Private FirePhase     As Double
Private CurWeaponPos  As D3DVECTOR
Private CurWeaponMtrx As D3DMATRIX
Private CurWeaponAddZ As Single
Private vDeclThing    As Direct3DVertexDeclaration
Private VSThing       As Direct3DVertexShader9
Private PSThing       As Direct3DPixelShader9
Public Event OnTake(ByVal sender As String, ByVal index As Long)
Public Sub Take(ByVal sender As String, ByVal index As Long)
   RaiseEvent OnTake(sender, index)
End Sub
Friend Sub Initialize()
   On Error Resume Next
   VSCreate
   PSCreate
   MeshCnt = MeshCnt + 1
   ReDim Preserve Mesh(0 To MeshCnt - 1)
   Set Mesh(MeshCnt - 1) = New clsMesh
   Mesh(MeshCnt - 1).Init App.Path & "\3DEngine\Meshes\medkit.mesh", 1, Vec3(0, 0, 0)
End Sub
Friend Sub Save(ByVal nf As Long)
   On Error Resume Next
   Dim i As Long
   Put #nf, , ThingCnt
   For i = 0 To ThingCnt - 1
      Put #nf, , Thing(i)
   Next i
End Sub
Friend Sub Load(ByVal nf As Long)
   On Error Resume Next
   Dim i As Long
   Get #nf, , ThingCnt
   For i = 0 To ThingCnt - 1
      Get #nf, , Thing(i)
   Next i
End Sub
Friend Sub Insert(ByRef t As tThing2D, ByVal meshname As String, ByVal texturename As String, ByVal pth As String)
   On Error Resume Next
   'ReDim Preserve tex(0 To ThingCnt - 1)
   'Set tex(ThingCnt - 1) = CreateTextureFromFileEx(Dev, App.Path & "\3DEngine\Meshes\" & texturename, 0, 0, 0, D3DUSAGE_NONE, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_BOX, D3DX_FILTER_LINEAR, 0)
   ReDim Preserve tex(0 To ThingCnt)
   Set tex(ThingCnt) = CreateTextureFromFileEx(Dev, App.Path & "\3DEngine\Meshes\" & texturename, 0, 0, 0, D3DUSAGE_NONE, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_BOX, D3DX_FILTER_LINEAR, 0)
   ReDim Preserve Thing(ThingCnt)
   ReDim Preserve ThingPath(ThingCnt)
   Thing(ThingCnt).InitPos = Vec3(t.x, t.y + 4, t.z)
   Thing(ThingCnt).InitPos.y = Thing(ThingCnt).InitPos.y + frmMain.mapClass.GetFloor(Thing(ThingCnt).InitPos)
   Thing(ThingCnt).Pos = Thing(ThingCnt).InitPos
   Thing(ThingCnt).MeshInd = 0 'ThingCnt - 1 'MeshCnt - 1
   ThingPath(ThingCnt) = pth
   ThingCnt = ThingCnt + 1
End Sub
Friend Sub AddPOV(Thing2D As tThing2D)
   On Error Resume Next
   Select Case Thing2D.tType
   Case 0
      StartPos = Vec3(Thing2D.x, Thing2D.y + 0.5, Thing2D.z)
      StartPos.y = frmMain.mapClass.GetFloor(StartPos) + 4 'height
   Case 1
      '            MeshCnt = MeshCnt + 1
      '            ReDim Mesh(0 To MeshCnt - 1)
      '            ReDim tex(0 To MeshCnt - 1)
      '            Set Mesh(0) = New clsMesh
      '            Mesh(0).Init App.Path & "\3DEngine\Meshes\raptor_01.mesh", 1, Vec3(0, 0, 0)
      '            Set tex(0) = CreateTextureFromFileEx(Dev, App.Path & "\3DEngine\Meshes\tiger.jpg", 0, 0, 0, D3DUSAGE_NONE, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_BOX, D3DX_FILTER_LINEAR, 0)
      '            ReDim Preserve Thing(ThingCnt)
      '            Thing(ThingCnt).InitPos = Vec3(Thing2D.x, Thing2D.y + 0.5, Thing2D.z)
      '            Thing(ThingCnt).InitPos.y = frmMain.mapClass.GetFloor(Thing(ThingCnt).InitPos) + 20
      '            Thing(ThingCnt).Pos = Thing(ThingCnt).InitPos
      '            Thing(ThingCnt).MeshInd = 0
      '            ThingCnt = ThingCnt + 1
   End Select
End Sub
'Private Function TakeThing(ByVal Ind As Long) As Boolean
'    On Error Resume Next
'    Select Case Thing(Ind).MeshInd
'        Case 0
'           ' Remove Ind
'            Take "medkit", Ind
'    End Select
'End Function
Friend Sub Remove(ByVal Ind As Long)
   On Error Resume Next
   Dim i As Long
   For i = Ind To ThingCnt - 1
      ThingPath(i) = ThingPath(i + 1) 'overwrite
      Thing(i) = Thing(i + 1)
   Next
   'ThingPath(Ind) = ThingPath(ThingCnt - 1) 'overwrite
   ' Thing(Ind) = Thing(ThingCnt - 1)
   ThingCnt = ThingCnt - 1
End Sub
Friend Sub Tick()
   On Error Resume Next
   Dim Mtrx As D3DMATRIX
   Dim i    As Long
   For i = 0 To ThingCnt - 1
      ' Thing(i).Pos = Vec3(Thing(i).InitPos.x, Thing(i).InitPos.y + Sin(NowTime * 5) * 0.5, Thing(i).InitPos.z) 'move up and down
      MatrixTranslation Mtrx, Thing(i).Pos.x, Thing(i).Pos.y, Thing(i).Pos.z 'get current matrix position
      MatrixRotationY Thing(i).Mtrx, 1 'NowTime * 1 '2.5  'spin speed 1=no spin
      MatrixMultiply Thing(i).Mtrx, Thing(i).Mtrx, Mtrx 'set new matrix position
      If PlIntersect(Thing(i).Pos) = True Then
         '            If TakeThing(i) = True Then
         '                Exit For
         '            End If
      End If
   Next i
End Sub
Friend Sub Draw()
   On Error Resume Next
   Dim Mtrx As D3DMATRIX
   Dim v3   As D3DVECTOR
   Dim i    As Long
   Dev.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE
   Dev.SetRenderState D3DRS_ZWRITEENABLE, D3D_TRUE
   Dev.SetRenderState D3DRS_ALPHABLENDENABLE, D3D_FALSE
   Dev.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
   Dev.SetVertexDeclaration vDeclThing
   Dev.SetVertexShader VSThing
   Dev.SetPixelShader PSThing
   For i = 0 To ThingCnt - 1
      MatrixTranspose Mtrx, Thing(i).Mtrx
      Vec3Subtract v3, PlPos, Thing(i).Pos
      Vec3TransformNormal VSConst.CamPos, v3, Mtrx
      Vec3TransformNormal VSConst.LitePos, SunLightPos, Mtrx
      MatrixMultiply Mtrx, Thing(i).Mtrx, mView
      MatrixMultiply Mtrx, Mtrx, mProj
      MatrixTranspose VSConst.WVP, Mtrx
      MatrixTranspose VSConst.World, Thing(i).Mtrx
      MatrixMultiply Mtrx, Thing(i).Mtrx, mViewShad
      MatrixMultiply Mtrx, Mtrx, mProjShad
      MatrixTranspose VSConst.mShad, Mtrx
      Dev.SetVertexShaderConstantF 0, VarPtr(VSConst), 16
      Dev.SetTexture 0, CreateTextureFromFileEx(Dev, App.Path & "\3DEngine\Meshes\icon" & i & ".jpg", 0, 0, 0, D3DUSAGE_NONE, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_BOX, D3DX_FILTER_LINEAR, 0) 'Thing(i).MeshInd
      Mesh(Thing(i).MeshInd).Draw
   Next i
End Sub
Friend Sub DrawShade()
   Dim Mtrx As D3DMATRIX
   Dim i    As Long
   Dev.SetVertexDeclaration vDeclThing
   For i = 0 To ThingCnt - 1
      MatrixMultiply Mtrx, Thing(i).Mtrx, mViewShad
      MatrixMultiply Mtrx, Mtrx, mProjShad
      MatrixTranspose VSConst.WVP, Mtrx
      Dev.SetVertexShaderConstantF 0, VarPtr(VSConst), 4
      Mesh(Thing(i).MeshInd).Draw
   Next i
End Sub
Friend Sub Fire()
   On Error Resume Next
   Dim RISV As D3DVECTOR
   Dim RISN As D3DVECTOR
   Dim v    As D3DVECTOR
   '
   Vec3Add v, PlDir, PlSpeed
   'frmMain.consoleClass.Display CStr(PlDir.x)
   '
   'Vec3Scale v, v, 0.35 ' default speed of .35
   '
   frmMain.ProjectileClass.add CurWeaponPos, v
   '        apiGetCursorPos pc
   '        mx = (pc.x - oldpoint.x) * MouseSensX
   '        my_ = (pc.y - oldpoint.y) * MouseSensY
   'PlAngle = PlAngle + mx
   'PlDiff = PlDiff + my_
   ' apiSetCursorPos ScrCenterX, ScrCenterY
   '    Dim mx  As Single
   '    Dim my_ As Single
   '    Dim pC  As POINTAPI
   '    Dim i   As Long
   'For i = 1 To 100
   '        apiGetCursorPos pc
   '        mx = (pc.x - oldpoint.x) * MouseSensX
   '        my_ = (pc.y - oldpoint.y) * MouseSensY
   '        PlAngle = PlAngle + mx
   '        PlDiff = PlDiff + my_
   '       ' apiSetCursorPos ScrCenterX, ScrCenterY
End Sub
Friend Sub SelectorTick()
   On Error Resume Next
   'Static os      As D3DVECTOR
   'Static VSThing As D3DVECTOR
   'Dim Mtrx       As D3DMATRIX
   Dim v As D3DVECTOR
   'Dim a          As Single
   ' Dim LL         As Single
   ' Dim RISV       As D3DVECTOR
   ' Dim RISN       As D3DVECTOR
   Vec3Scale v, PlRight, 0 '0.2 + 0.2 * 1 'calculate offset to the right side arm, 0=no offset
   '
   Vec3Add CurWeaponPos, PlPos, v 'set new offset
   '
   Vec3Scale v, PlUp, -0.55 'calculate offset below, -0.55=fire from below at waist level
   '
   Vec3Add CurWeaponPos, CurWeaponPos, v 'set new offset
End Sub
Private Sub VSCreate()
   On Error Resume Next
   Dim vsDecl(3) As D3DVERTEXELEMENT9
   vsDecl(0) = VertexElement(0, 0, D3DDECLTYPE_FLOAT3, D3DDECLMETHOD_DEFAULT, D3DDECLUSAGE_POSITION, 0)
   vsDecl(1) = VertexElement(0, 12, D3DDECLTYPE_FLOAT3, D3DDECLMETHOD_DEFAULT, D3DDECLUSAGE_NORMAL, 0)
   vsDecl(2) = VertexElement(0, 24, D3DDECLTYPE_FLOAT2, D3DDECLMETHOD_DEFAULT, D3DDECLUSAGE_TEXCOORD, 0)
   vsDecl(3) = VertexElementEnd
   Set vDeclThing = Dev.CreateVertexDeclaration(VarPtr(vsDecl(0)))
   Set VSThing = Dev.CreateVertexShaderFromFile(App.Path & "\3DEngine\ShaderCompiler\Out\Model.vsh.shader")
End Sub
Private Sub PSCreate()
   On Error Resume Next
   Set PSThing = Dev.CreatePixelShaderFromFile(App.Path & "\3DEngine\ShaderCompiler\Out\Model.psh.shader")
End Sub
Friend Sub ThingsTerminate()
   On Error Resume Next
   Dim i As Long
   For i = 0 To MeshCnt - 1
      Set tex(i) = Nothing
      Set Mesh(i) = Nothing
   Next i
   Set vDeclThing = Nothing
   Set VSThing = Nothing
   Set PSThing = Nothing
End Sub
