VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type MapSubset
    vBuf As Direct3DVertexBuffer9
    iBuf As Direct3DIndexBuffer9
    tCnt As Long
    Vert3() As Map3Vertex
    vCnt As Long
    Ind() As Integer
    iCnt As Long
    Enabled As Boolean
End Type
Private vDecl      As Direct3DVertexDeclaration
Private VS         As Direct3DVertexShader9
Private ps         As Direct3DPixelShader9
Private VSZ        As Direct3DVertexShader9
Private psz        As Direct3DPixelShader9
Private vSize      As Long
Private MapSS()    As MapSubset
Private TexC()     As Direct3DTexture9
Private TexN()     As Direct3DTexture9
Private TexLS1     As Direct3DTexture9
Private TexLS2     As Direct3DTexture9
Private TexLSMask  As Direct3DTexture9
Private TexSkyUp   As Direct3DTexture9
Private TexSkyDown As Direct3DTexture9
Private TexSun     As Direct3DTexture9
Private TexButC    As Direct3DTexture9
Private TexButL    As Direct3DTexture9
Private Sky        As clsSky
Public Event OnLoadComplete()
Public Sub LoadComplete()
    On Error Resume Next
    RaiseEvent OnLoadComplete
End Sub

Private Sub AddFace(Vert3() As Map3Vertex, ByVal vCnt As Long, Ind() As Integer, ByVal iCnt As Long, ByVal m As Long, ByVal dInd As Long)
    On Error Resume Next
    Dim pVert As Long
    Dim pInd  As Long
    Dim i     As Long
    If m > 0 Then
        pVert = MapSS(m, dInd).vCnt
        MapSS(m, dInd).vCnt = MapSS(m, dInd).vCnt + vCnt
        ReDim Preserve MapSS(m, dInd).Vert3(MapSS(m, dInd).vCnt - 1)
        For i = 0 To vCnt - 1
            MapSS(m, dInd).Vert3(pVert + i) = Vert3(i)
        Next i
        pInd = MapSS(m, dInd).iCnt
        MapSS(m, dInd).iCnt = MapSS(m, dInd).iCnt + iCnt
        ReDim Preserve MapSS(m, dInd).Ind(MapSS(m, dInd).iCnt - 1)
        For i = 0 To iCnt - 1
            MapSS(m, dInd).Ind(pInd + i) = Ind(i) + pVert
        Next i
    End If
End Sub
Friend Sub Load(fName As String)
    On Error Resume Next
    Dim GMapName As String
    Dim nf       As Long
    Dim s        As String
    Dim s1       As String
    Dim i        As Long
    Dim v4       As D3DVECTOR4
    ReDim TexScale(9)
    ReDim TexAlign(9)
    For i = 1 To 9
        TexScale(i) = 1 / 8
    Next i
    ReDim TexC(1 To 9)
    ReDim TexN(1 To 9)
    nf = FreeFile
    Open fName For Input As #nf
    While Not EOF(nf)
        Line Input #nf, s
        s = LCase(Trim(s))
        s1 = StrCut(s)
        If left(s1, 4) = "texc" Then
            i = Val(Mid(s1, 5))
            s1 = "texc"
        ElseIf left(s1, 4) = "texn" Then
            i = Val(Mid(s1, 5))
            s1 = "texn"
        ElseIf left(s1, 3) = "tex" Then
            i = Val(Mid(s1, 4))
            s1 = "tex"
        End If
        Select Case s1
            Case "gmap"
                GMapName = App.Path & "\3DEngine\maps\" & s
            Case "hmap"
                Set LS = New clsLandScape
                LS.Init App.Path & "\3DEngine\maps\" & s, 2, 35
            Case "lstex1"
                Set TexLS1 = CreateTextureFromFileEx(Dev, App.Path & "\3DEngine\Maptex\" & s, 0, 0, 0, D3DUSAGE_NONE, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_BOX, D3DX_FILTER_LINEAR, 0)
            Case "lstex2"
                Set TexLS2 = CreateTextureFromFileEx(Dev, App.Path & "\3DEngine\Maptex\" & s, 0, 0, 0, D3DUSAGE_NONE, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_BOX, D3DX_FILTER_LINEAR, 0)
            Case "lsmask"
                Set TexLSMask = CreateTextureFromFileEx(Dev, App.Path & "\3DEngine\Maptex\" & s, 0, 0, 0, D3DUSAGE_NONE, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_BOX, D3DX_FILTER_LINEAR, 0)
            Case "texc"
                Set TexC(i) = CreateTextureFromFileEx(Dev, App.Path & "\3DEngine\Maptex\" & s, 0, 0, 0, D3DUSAGE_NONE, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_BOX, D3DX_FILTER_LINEAR, 0)
            Case "texn"
                Set TexN(i) = CreateTextureFromFileEx(Dev, App.Path & "\3DEngine\Maptex\" & s, 0, 0, 0, D3DUSAGE_NONE, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_BOX, D3DX_FILTER_LINEAR, 0)
            Case "tex"
                TexScale(i) = Val(s) * 0.01 / 8
                If right(s, 1) = "a" Then TexAlign(i) = 1 Else TexAlign(i) = 0
            Case "skytexup"
                Set TexSkyUp = CreateTextureFromFileEx(Dev, App.Path & "\3DEngine\Maptex\" & s, 0, 0, 0, D3DUSAGE_NONE, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_BOX, D3DX_FILTER_LINEAR, 0)
            Case "skytexdown"
                Set TexSkyDown = CreateTextureFromFileEx(Dev, App.Path & "\3DEngine\Maptex\" & s, 0, 0, 0, D3DUSAGE_NONE, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_BOX, D3DX_FILTER_LINEAR, 0)
            Case "suntex"
                Set TexSun = CreateTextureFromFileEx(Dev, App.Path & "\3DEngine\Maptex\" & s, 0, 0, 0, D3DUSAGE_NONE, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_BOX, D3DX_FILTER_LINEAR, 0)
            Case "buttex"
                Set TexButC = CreateTextureFromFileEx(Dev, App.Path & "\3DEngine\Maptex\" & s, 0, 0, 0, D3DUSAGE_NONE, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_BOX, D3DX_FILTER_LINEAR, 0)
            Case "butltex"
                Set TexButL = CreateTextureFromFileEx(Dev, App.Path & "\3DEngine\Maptex\" & s, 0, 0, 0, D3DUSAGE_NONE, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_BOX, D3DX_FILTER_LINEAR, 0)
            Case "sunpos"
                SunLightPos.x = Val(StrCut(s))
                SunLightPos.y = Val(StrCut(s))
                SunLightPos.z = Val(s)
            Case "ambient"
                PSConst.Ambient.x = Val(StrCut(s))
                PSConst.Ambient.y = Val(StrCut(s))
                PSConst.Ambient.z = Val(s)
            Case "fogval"
                PSConst.FogVal = Vec4(255, 1, 0, Val(s))
            Case "fogcolor"
                PSConst.FogColor.x = Val(StrCut(s))
                PSConst.FogColor.y = Val(StrCut(s))
                PSConst.FogColor.z = Val(s)
            Case "diffuse"
                VSConst.Diffuse.r = Val(StrCut(s))
                VSConst.Diffuse.G = Val(StrCut(s))
                VSConst.Diffuse.b = Val(s)
                VSConst.Diffuse.a = 0.25
            Case "specular"
                VSConst.Specular.r = Val(StrCut(s))
                VSConst.Specular.G = Val(StrCut(s))
                VSConst.Specular.b = Val(s)
                VSConst.Specular.a = 0.5
        End Select
    Wend
    Close #nf
    GMapLoad GMapName
    LoadComplete
End Sub
Private Function StrCut(s As String) As String
    On Error Resume Next
    Dim i As Long
    i = InStr(s, " ")
    StrCut = Trim(left(s, i))
    s = Trim(Mid(s, i + 1))
End Function
Private Sub GMapLoad(fName As String)
    On Error Resume Next
    Dim Vert3() As Map3Vertex
    Dim vCnt    As Long
    Dim Ind()   As Integer
    Dim iCnt    As Long
    Dim i       As Long
    Dim di      As Long
    Dim m       As Long
    Dim n       As Long
    Dim t       As Long
    Dim nf      As Long
    Dim s       As String
    Dim Thing2D As tThing2D
    nf = FreeFile
    Open fName For Binary As #nf
    Get #nf, , SectorCnt
    ReDim Sector(SectorCnt - 1)
    For i = 0 To SectorCnt - 1
        Set Sector(i) = New clsSector
        Sector(i).Load nf
    Next i
    Get #nf, , t
    For i = 0 To t - 1
        Get #nf, , Thing2D
        frmMain.itemclass.AddPOV Thing2D
    Next i
    Close #nf
    ReDim MapSS(1 To 9, 0)
    For i = 0 To SectorCnt - 1
        If Sector(i).Tag > 0 Then
            DoorCnt = DoorCnt + 1
            di = DoorCnt
            ReDim Preserve MapSS(1 To 9, DoorCnt)
            ReDim Preserve Door(1 To DoorCnt)
            MatrixIdentity Door(DoorCnt).Mtrx
            Door(DoorCnt).yMin = Sector(i).yMin
            Door(DoorCnt).yMax = Sector(i).yMax
            Door(DoorCnt).Tag = Sector(i).Tag
            If Sector(i).Action > 0 Then
                Door(DoorCnt).pMin = 0
                Door(DoorCnt).pMax = Sector(i).Action
            Else
                Door(DoorCnt).pMin = Sector(i).Action
                Door(DoorCnt).pMax = 0
            End If
            Door(DoorCnt).Sect = i
            Door(DoorCnt).Pos.x = (Sector(i).xMin + Sector(i).xMax) * 0.5
            Door(DoorCnt).Pos.z = (Sector(i).zMin + Sector(i).zMax) * 0.5
        Else
            di = 0
        End If
        Sector(i).GetFaceUp Vert3(), Ind(), vCnt, iCnt, m
        AddFace Vert3(), vCnt, Ind(), iCnt, m, di
        Sector(i).GetFaceDown Vert3(), Ind(), vCnt, iCnt, m
        AddFace Vert3(), vCnt, Ind(), iCnt, m, di
        For n = 0 To Sector(i).GetVCnt - 1
            Sector(i).GetFace n, Vert3(), Ind(), vCnt, iCnt, m, t
            If t > 0 Then frmMain.buttonClass.add Vert3(), t, i
            AddFace Vert3(), vCnt, Ind(), iCnt, m, di
        Next n
    Next i
    vSize = Len(Vert3(0))
    For n = 0 To DoorCnt
        For m = 1 To 9
            If MapSS(m, n).vCnt > 0 Then
                MapSS(m, n).Enabled = True
                MapSS(m, n).tCnt = MapSS(m, n).iCnt \ 3
                Set MapSS(m, n).vBuf = Dev.CreateVertexBuffer(MapSS(m, n).vCnt * vSize, D3DUSAGE_NONE, D3DFVF_XYZ, D3DPOOL_MANAGED)
                MapSS(m, n).vBuf.SetData 0, MapSS(m, n).vCnt * vSize, VarPtr(MapSS(m, n).Vert3(0)), 0
                ReDim MapSS(m, n).Vert3(0)
                Set MapSS(m, n).iBuf = Dev.CreateIndexBuffer(MapSS(m, n).iCnt * Len(Ind(0)), D3DUSAGE_NONE, D3DFMT_INDEX16, D3DPOOL_MANAGED)
                MapSS(m, n).iBuf.SetData 0, MapSS(m, n).iCnt * Len(Ind(0)), VarPtr(MapSS(m, n).Ind(0)), 0
                ReDim MapSS(m, n).Ind(0)
            End If
        Next m
    Next n
    VSCreate
    PSCreate
    Set Sky = New clsSky
End Sub
Friend Function PointInMap(p As D3DVECTOR) As Boolean
    Dim i As Long
    For i = 0 To SectorCnt - 1
        If Sector(i).PointInSector(p) Then PointInMap = True: Exit Function
    Next i
End Function
Friend Function GetFloor(p As D3DVECTOR) As Single
    Dim h As Single
    Dim i As Long
    LS.GetHeight p.x, p.z, h
    If h > p.y Then h = -999999
    For i = 0 To SectorCnt - 1
        Sector(i).GefFloor p, h
    Next i
    GetFloor = h
End Function
Friend Function SphereInMap(p As D3DVECTOR, ByVal r As Single, OutN As D3DVECTOR, Optional OutD As Single) As Boolean
    On Error Resume Next
    Dim i As Long
    Dim h As Single
    LS.GetHeight p.x, p.z, h
    If p.y < h + r Then
        LS.GetNorm p.x, p.z, OutN
        SphereInMap = True
        OutD = 0.33
        Exit Function
    End If
    For i = 0 To SectorCnt - 1
        If Sector(i).SphereInSector(p, r, OutN) Then
            SphereInMap = True
            OutD = 0.66
            Exit Function
        End If
    Next i
End Function
Friend Function RayInMap(p As D3DVECTOR, d As D3DVECTOR, OutP As D3DVECTOR, OutN As D3DVECTOR, OutL As Single) As Boolean
    On Error Resume Next
    Dim i As Long, v As D3DVECTOR, vN As D3DVECTOR, vv As D3DVECTOR, LL As Single, f As Boolean
    If LS.RayInterSector(p, d, v, vN) Then
        Vec3Subtract vv, v, p
        OutL = Vec3LengthSq(vv)
        OutP = v
        OutN = vN
        f = True
    Else
        OutL = 999999999
    End If
    For i = 0 To SectorCnt - 1
        If Sector(i).RayInSector(p, d, v, vN) Then
            Vec3Subtract vv, v, p
            LL = Vec3LengthSq(vv)
            If OutL > LL Then
                OutL = LL
                OutP = v
                OutN = vN
                f = True
            End If
        End If
    Next i
    If f Then
        OutL = Sqr(OutL)
    Else
        Vec3Scale v, d, 10000
        Vec3Add OutP, p, v
    End If
    RayInMap = f
End Function
Friend Sub Draw()
    On Error Resume Next
    Dim Mtrx As D3DMATRIX
    Dim m    As Long
    Dim n    As Long
    Sky.Draw PlPos, TexSkyUp, TexSkyDown, TexSun
    VSConst.CamPos = PlPos
    VSConst.LitePos = SunLightPos
    MatrixMultiply Mtrx, mView, mProj
    MatrixTranspose VSConst.WVP, Mtrx
    MatrixMultiply Mtrx, mViewShad, mProjShad
    MatrixTranspose VSConst.mShad, Mtrx
    Dev.SetVertexShaderConstantF 0, VarPtr(VSConst), 16
    Dev.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE
    Dev.SetRenderState D3DRS_ALPHABLENDENABLE, D3D_FALSE
    Dev.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    Dev.SetVertexDeclaration vDecl
    If ZOptimize Then
        Dev.SetVertexShader VSZ
        Dev.SetPixelShader psz
        For m = 1 To 9
            If MapSS(m, 0).Enabled Then
                Dev.SetStreamSource 0, MapSS(m, 0).vBuf, 0, vSize
                Dev.SetIndices MapSS(m, 0).iBuf
                Dev.DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, 0, MapSS(m, 0).vCnt, 0, MapSS(m, 0).tCnt
            End If
        Next m
        For n = 1 To DoorCnt
            MatrixMultiply Mtrx, Door(n).Mtrx, mView
            MatrixMultiply Mtrx, Mtrx, mProj
            MatrixTranspose VSConst.WVP, Mtrx
            MatrixMultiply Mtrx, Door(n).Mtrx, mViewShad
            MatrixMultiply Mtrx, Mtrx, mProjShad
            MatrixTranspose VSConst.mShad, Mtrx
            Dev.SetVertexShaderConstantF 0, VarPtr(VSConst), 16
            For m = 1 To 9
                If MapSS(m, n).Enabled Then
                    Dev.SetStreamSource 0, MapSS(m, n).vBuf, 0, vSize
                    Dev.SetIndices MapSS(m, n).iBuf
                    Dev.DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, 0, MapSS(m, n).vCnt, 0, MapSS(m, n).tCnt
                End If
            Next m
        Next n
        Dev.SetRenderState D3DRS_ZFUNC, D3DCMP_EQUAL
        Dev.SetRenderState D3DRS_ZWRITEENABLE, D3D_FALSE
        MatrixMultiply Mtrx, mView, mProj
        MatrixTranspose VSConst.WVP, Mtrx
        MatrixMultiply Mtrx, mViewShad, mProjShad
        MatrixTranspose VSConst.mShad, Mtrx
        Dev.SetVertexShaderConstantF 0, VarPtr(VSConst), 16
    End If
    Dev.SetVertexShader VS
    Dev.SetPixelShader ps
    Dev.SetTexture 3, RTTex
    For m = 1 To 9
        If MapSS(m, 0).Enabled Then
            Dev.SetTexture 0, TexC(m)
            Dev.SetTexture 1, TexN(m)
            Dev.SetStreamSource 0, MapSS(m, 0).vBuf, 0, vSize
            Dev.SetIndices MapSS(m, 0).iBuf
            Dev.DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, 0, MapSS(m, 0).vCnt, 0, MapSS(m, 0).tCnt
        End If
    Next m
    For n = 1 To DoorCnt
        MatrixMultiply Mtrx, Door(n).Mtrx, mView
        MatrixMultiply Mtrx, Mtrx, mProj
        MatrixTranspose VSConst.WVP, Mtrx
        MatrixMultiply Mtrx, Door(n).Mtrx, mViewShad
        MatrixMultiply Mtrx, Mtrx, mProjShad
        MatrixTranspose VSConst.mShad, Mtrx
        Dev.SetVertexShaderConstantF 0, VarPtr(VSConst), 16
        For m = 1 To 9
            If MapSS(m, n).Enabled Then
                Dev.SetTexture 0, TexC(m)
                Dev.SetTexture 1, TexN(m)
                Dev.SetStreamSource 0, MapSS(m, n).vBuf, 0, vSize
                Dev.SetIndices MapSS(m, n).iBuf
                Dev.DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, 0, MapSS(m, n).vCnt, 0, MapSS(m, n).tCnt
            End If
        Next m
    Next n
    If ZOptimize Then
        Dev.SetRenderState D3DRS_ZFUNC, D3DCMP_LESS
        Dev.SetRenderState D3DRS_ZWRITEENABLE, D3D_TRUE
    End If
    Dev.SetTexture 0, TexButC
    Dev.SetTexture 1, TexButL
    Dev.SetTexture 3, RTTex
    frmMain.buttonClass.Draw
    Dev.SetTexture 0, TexLS1
    Dev.SetTexture 1, TexLS2
    Dev.SetTexture 2, TexLSMask
    Dev.SetTexture 3, RTTex
    LS.Draw
End Sub
Friend Sub DrawShade()
    On Error Resume Next
    Dim Mtrx As D3DMATRIX
    Dim m    As Long
    Dim n    As Long
    MatrixMultiply Mtrx, mViewShad, mProjShad
    MatrixTranspose VSConst.WVP, Mtrx
    Dev.SetVertexShaderConstantF 0, VarPtr(VSConst), 4
    Dev.SetVertexDeclaration vDecl
    For m = 1 To 9
        If MapSS(m, 0).Enabled Then
            Dev.SetStreamSource 0, MapSS(m, 0).vBuf, 0, vSize
            Dev.SetIndices MapSS(m, 0).iBuf
            Dev.DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, 0, MapSS(m, 0).vCnt, 0, MapSS(m, 0).tCnt
        End If
    Next m
    For n = 1 To DoorCnt
        MatrixMultiply Mtrx, Door(n).Mtrx, mViewShad
        MatrixMultiply Mtrx, Mtrx, mProjShad
        MatrixTranspose VSConst.WVP, Mtrx
        Dev.SetVertexShaderConstantF 0, VarPtr(VSConst), 4
        For m = 1 To 9
            If MapSS(m, n).Enabled Then
                Dev.SetStreamSource 0, MapSS(m, n).vBuf, 0, vSize
                Dev.SetIndices MapSS(m, n).iBuf
                Dev.DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, 0, MapSS(m, n).vCnt, 0, MapSS(m, n).tCnt
            End If
        Next m
    Next n
End Sub
Private Sub VSCreate()
    On Error Resume Next
    Dim vsDecl(5) As D3DVERTEXELEMENT9
    vsDecl(0) = VertexElement(0, 0, D3DDECLTYPE_FLOAT3, D3DDECLMETHOD_DEFAULT, D3DDECLUSAGE_POSITION, 0)
    vsDecl(1) = VertexElement(0, 12, D3DDECLTYPE_FLOAT2, D3DDECLMETHOD_DEFAULT, D3DDECLUSAGE_TEXCOORD, 0)
    vsDecl(2) = VertexElement(0, 20, D3DDECLTYPE_FLOAT3, D3DDECLMETHOD_DEFAULT, D3DDECLUSAGE_TANGENT, 0)
    vsDecl(3) = VertexElement(0, 32, D3DDECLTYPE_FLOAT3, D3DDECLMETHOD_DEFAULT, D3DDECLUSAGE_BINORMAL, 0)
    vsDecl(4) = VertexElement(0, 44, D3DDECLTYPE_FLOAT3, D3DDECLMETHOD_DEFAULT, D3DDECLUSAGE_NORMAL, 0)
    vsDecl(5) = VertexElementEnd
    Set vDecl = Dev.CreateVertexDeclaration(VarPtr(vsDecl(0)))
    Set VS = Dev.CreateVertexShaderFromFile(App.Path & "\3DEngine\ShaderCompiler\Out\Map.vsh.shader")
    Set VSZ = Dev.CreateVertexShaderFromFile(App.Path & "\3DEngine\ShaderCompiler\Out\MapZ.vsh.shader")
End Sub
Private Sub PSCreate()
    On Error Resume Next
    Set ps = Dev.CreatePixelShaderFromFile(App.Path & "\3DEngine\ShaderCompiler\Out\Map.psh.shader")
    Set psz = Dev.CreatePixelShaderFromFile(App.Path & "\3DEngine\ShaderCompiler\Out\MapZ.psh.shader")
End Sub
Friend Sub Terminate()
    On Error Resume Next
    Dim i As Long, j As Long
    Set LS = Nothing
    frmMain.buttonClass.Terminate
    For i = 0 To SectorCnt - 1
        Set Sector(i) = Nothing
    Next i
    For j = 0 To DoorCnt
        For i = 1 To 9
            Set MapSS(i, j).vBuf = Nothing
            Set MapSS(i, j).iBuf = Nothing
        Next i
    Next j
    For i = 1 To 9
        Set TexC(i) = Nothing
        Set TexN(i) = Nothing
    Next i
    Set TexLS1 = Nothing
    Set TexLS2 = Nothing
    Set TexLSMask = Nothing
    Set TexSkyUp = Nothing
    Set TexSkyDown = Nothing
    Set TexSun = Nothing
    Set TexButC = Nothing
    Set TexButL = Nothing
    Set VS = Nothing
    Set ps = Nothing
    Set VSZ = Nothing
    Set psz = Nothing
    Set vDecl = Nothing
    Set Sky = Nothing
End Sub
