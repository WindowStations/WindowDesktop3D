VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type Map2Vertex
    x As Single
    z As Single
    Mtrl As Integer
    Target As Integer
End Type
Dim Vert2() As Map2Vertex
Dim vCnt    As Long
Dim vInd()  As Integer
Dim iCnt    As Long
Dim MtrlUp  As Long, MtrlDown As Long
Public xMax As Single, xMin As Single
Public zMax As Single, zMin As Single
Public yMax As Single, yMin As Single, yMid As Single
Public Tag  As Long, Action As Long

Public Sub Load(ByVal nf As Long)
    Dim i As Long, d As Integer
    Get #nf, , d
    Tag = d
    Get #nf, , d
    Action = d
    Get #nf, , d
    yMax = d
    Get #nf, , d
    yMin = d
    Get #nf, , d
    MtrlUp = d
    Get #nf, , d
    MtrlDown = d
    Get #nf, , vCnt
    ReDim Vert2(vCnt - 1)
    For i = 0 To vCnt - 1
        Get #nf, , d
        Vert2(i).x = d
        Get #nf, , d
        Vert2(i).z = d
        Get #nf, , Vert2(i).Mtrl
        Get #nf, , Vert2(i).Target
    Next i
    yMid = (yMin + yMax) * 0.5
    xMin = 1000000
    xMax = -1000000
    zMin = 1000000
    zMax = -1000000
    For i = 0 To vCnt - 1
        If xMin > Vert2(i).x Then xMin = Vert2(i).x
        If xMax < Vert2(i).x Then xMax = Vert2(i).x
        If zMin > Vert2(i).z Then zMin = Vert2(i).z
        If zMax < Vert2(i).z Then zMax = Vert2(i).z
    Next i
    Triangulator
End Sub

Friend Function SphereInSector(p As D3DVECTOR, ByVal r As Single, OutN As D3DVECTOR) As Boolean
    Dim n  As Long, rr As Single, k As Single
    Dim v1 As D3DVECTOR2, v2 As D3DVECTOR2, v3 As D3DVECTOR2
    If p.x + r <= xMin Then Exit Function
    If p.x - r >= xMax Then Exit Function
    If p.z + r <= zMin Then Exit Function
    If p.z - r >= zMax Then Exit Function
    If p.y + r <= yMin Then
        Exit Function
    ElseIf p.y - r >= yMax Then
        Exit Function
    ElseIf p.y < yMin Then
        If PointInSector(Vec3(p.x, yMin, p.z)) Then
            OutN = Vec3(0, -1, 0)
            SphereInSector = True: Exit Function
        Else
            rr = Sqr(r * r - (yMin - p.y) * (yMin - p.y))
            OutN.y = (p.y - yMin) / r
            k = Sqr(1 - OutN.y * OutN.y)
        End If
    ElseIf p.y > yMax Then
        If PointInSector(Vec3(p.x, yMax, p.z)) Then
            OutN = Vec3(0, 1, 0)
            SphereInSector = True: Exit Function
        Else
            rr = Sqr(r * r - (p.y - yMax) * (p.y - yMax))
            OutN.y = (p.y - yMax) / r
            k = Sqr(1 - OutN.y * OutN.y)
        End If
    Else
        rr = r
        OutN.y = 0
        k = 1
    End If
    For n = 0 To vCnt - 1
        v1.x = p.x - Vert2(n).x
        v1.y = p.z - Vert2(n).z
        v2.x = InVert(n + 1).x - Vert2(n).x
        v2.y = InVert(n + 1).z - Vert2(n).z
        v3.x = InVert(n + 1).x - p.x
        v3.y = InVert(n + 1).z - p.z
        If Vec2Dot(v1, v2) < 0 Then
            If Vec2LengthSq(v1) < rr * rr Then
                Vec2Normalize v1, v1
                OutN.x = v1.x * k
                OutN.z = v1.y * k
                SphereInSector = True: Exit Function
            End If
        ElseIf Vec2Dot(v3, v2) < 0 Then
            If Vec2LengthSq(v3) < rr * rr Then
                Vec2Normalize v3, v3
                OutN.x = v3.x * -k
                OutN.z = v3.y * -k
                SphereInSector = True: Exit Function
            End If
        Else
            Vec2Normalize v2, v2
            If Abs(Vec2CCW(v2, v1)) < rr Then
                OutN.x = v2.y * k
                OutN.z = v2.x * -k
                SphereInSector = True: Exit Function
            End If
        End If
    Next n
End Function

Friend Function RayInSector(p As D3DVECTOR, d As D3DVECTOR, OutP As D3DVECTOR, OutN As D3DVECTOR) As Boolean
    Dim n   As Long, c As Boolean
    Dim v   As D3DVECTOR, vv As D3DVECTOR2
    Dim v1  As D3DVECTOR, v2 As D3DVECTOR
    Dim iv1 As Map2Vertex, iv2 As Map2Vertex
    Dim s   As Single, ts As Single
    If (p.x <= xMin And d.x <= 0) Or (p.x >= xMax And d.x >= 0) Or (p.y <= yMin And d.y <= 0) Or (p.y >= yMax And d.y >= 0) Or _
       (p.z <= zMin And d.z <= 0) Or (p.z >= zMax And d.z >= 0) Then Exit Function
    Select Case p.y
        Case Is > yMax
            Vec3Scale v1, d, (yMax - p.y) / d.y
            Vec3Add v1, v1, p
            If PointInSector(Vec3(v1.x, (yMin + yMax) * 0.5, v1.z)) Then
                OutP = v1
                OutN = Vec3(0, 1, 0)
                RayInSector = True
                Exit Function
            End If
            Vec3Scale v2, d, (yMin - p.y) / d.y
            Vec3Add v2, v2, p
        Case Is < yMin
            Vec3Scale v1, d, (yMin - p.y) / d.y
            Vec3Add v1, v1, p
            If PointInSector(Vec3(v1.x, (yMin + yMax) * 0.5, v1.z)) Then
                OutP = v1
                OutN = Vec3(0, -1, 0)
                RayInSector = True
                Exit Function
            End If
            Vec3Scale v2, d, (yMax - p.y) / d.y
            Vec3Add v2, v2, p
        Case Else
            v1 = p
            Vec3Scale v, d, 1000
            Vec3Add v2, v1, v
            If v2.y > yMax Then
                Vec3Scale v, v, (yMax - v1.y) / (v2.y - v1.y)
                Vec3Add v2, v1, v
            ElseIf v2.y < yMin Then
                Vec3Scale v, v, (yMin - v1.y) / (v2.y - v1.y)
                Vec3Add v2, v1, v
            End If
    End Select
    iv1.x = v1.x: iv1.z = v1.z
    iv2.x = v2.x: iv2.z = v2.z
    s = 1000000
    For n = 0 To vCnt - 1
        If Intersect(Vert2(n), InVert(n + 1), iv1, iv2) Then
            ts = LenToIS(Vert2(n), InVert(n + 1), v1, v2, v)
            If s > ts Then
                s = ts
                OutP = v
                vv = Subt(Vert2(n), InVert(n + 1))
                v.x = vv.x: v.y = 0: v.z = vv.y
                Vec3Cross v, v, Vec3(0, 1, 0)
                Vec3Normalize OutN, v
                RayInSector = True
            End If
        End If
    Next n
End Function

Private Function LenToIS(v1 As Map2Vertex, v2 As Map2Vertex, v3 As D3DVECTOR, v4 As D3DVECTOR, OutV As D3DVECTOR) As Single
    Dim k1 As Single, k2 As Single
    Dim v  As D3DVECTOR, iv As Map2Vertex, vv2 As D3DVECTOR2
    vv2.x = v3.x - v1.x
    vv2.y = v3.z - v1.z
    k1 = CCW(Subt(v2, v1), vv2)
    vv2.x = v1.x - v4.x
    vv2.y = v1.z - v4.z
    k2 = CCW(Subt(v2, v1), vv2)
    k1 = k1 / (k1 + k2)
    Vec3Subtract v, v4, v3
    Vec3Scale v, v, k1
    Vec3Add v, v, v3
    OutV = v
    LenToIS = k1
End Function

Friend Function PlayerInSector(p As D3DVECTOR, OldP As D3DVECTOR, ByVal r As Single, ByVal HD As Single, ByVal HU As Single, ByVal HF As Single, _
   OutN As D3DVECTOR, OutFloor As Single) As Boolean
    Dim n  As Long, rr As Single, f As Boolean
    Dim v1 As D3DVECTOR2, v2 As D3DVECTOR2, v3 As D3DVECTOR2
    OutFloor = -999999
    If (p.x + r <= xMin) Or (p.x - r >= xMax) Or (p.z + r <= zMin) Or (p.z - r >= zMax) Or (p.y + HU <= yMin) Or (p.y >= yMax + HF) Then Exit Function
    If PointInSector(Vec3(p.x, yMid, p.z)) Then
        If p.y >= yMax + HD Then
            OutFloor = yMax
            Exit Function
        ElseIf p.y > yMid Then
            OutN = Vec3(0, 1, 0)
            OutFloor = yMax
        Else
            OutN = Vec3(0, -1, 0)
        End If
        PlayerInSector = True
        Exit Function
    End If
    Do
        f = True
        For n = 0 To vCnt - 1
            v1.x = p.x - Vert2(n).x
            v1.y = p.z - Vert2(n).z
            v2.x = InVert(n + 1).x - Vert2(n).x
            v2.y = InVert(n + 1).z - Vert2(n).z
            v3.x = InVert(n + 1).x - p.x
            v3.y = InVert(n + 1).z - p.z
            Vec2Normalize v2, v2
            rr = Vec2CCW(v1, v2)
            If (Vec2Dot(v1, v2) > 0) And (Vec2Dot(v3, v2) > 0) And (rr > 0) And (rr < r) Then
                If p.y >= yMax + HD Then
                    If rr < r * 0.99 Then
                        If OutFloor < yMax Then OutFloor = yMax
                        f = False
                    End If
                Else
                    OutN.x = v2.y
                    OutN.z = -v2.x
                    OutN.y = 0
                    Exit Do
                End If
            End If
        Next n
        For n = 0 To vCnt - 1
            v1.x = p.x - Vert2(n).x
            v1.y = p.z - Vert2(n).z
            rr = Vec2Length(v1)
            If rr < r Then
                If p.y >= yMax + HD Then
                    If rr < r * 0.99 Then
                        If OutFloor < yMax Then OutFloor = yMax
                    End If
                Else
                    v1.x = OldP.x - Vert2(n).x
                    v1.y = OldP.z - Vert2(n).z
                    Vec2Normalize v1, v1
                    OutN.x = v1.x
                    OutN.z = v1.y
                    OutN.y = 0
                    Exit Do
                End If
            End If
        Next n
        Exit Function
    Loop
    If p.y < yMin Then
        If p.y + HU - yMin < (r - rr) Then
            OutN.y = -1
            OutN.x = 0
            OutN.z = 0
        End If
    ElseIf p.y > yMax Then
        If yMax - (p.y - HD) < (r - rr) Then
            OutN.y = 1
            OutN.x = 0
            OutN.z = 0
        End If
    End If
    PlayerInSector = True
End Function

Private Function InVert(ByVal i As Long) As Map2Vertex
    InVert = Vert2((i + vCnt) Mod vCnt)
End Function

Friend Sub GefFloor(p As D3DVECTOR, CurFloor As Single)
    Dim n  As Long, c As Boolean
    Dim X1 As Single, z1 As Single, X2 As Single, z2 As Single
    If (p.x < xMin) Or (p.x > xMax) Or (p.z < zMin) Or (p.z > zMax) Or (p.y < yMax) Then Exit Sub
    For n = 0 To vCnt - 1
        X1 = Vert2(n).x
        z1 = Vert2(n).z
        X2 = InVert(n + 1).x
        z2 = InVert(n + 1).z
        If (z1 > p.z) Xor (z2 > p.z) Then
            If p.x > X1 + (X2 - X1) * (p.z - z1) / (z2 - z1) Then c = Not c
        End If
    Next n
    If c Then
        If CurFloor < yMax Then CurFloor = yMax
    End If
End Sub

Friend Function PointInSector(p As D3DVECTOR) As Boolean
    Dim n  As Long, c As Boolean
    Dim X1 As Single, z1 As Single, X2 As Single, z2 As Single
    If (p.x < xMin) Or (p.x > xMax) Or (p.y < yMin) Or (p.y > yMax) Or (p.z < zMin) Or (p.z > zMax) Then Exit Function
    For n = 0 To vCnt - 1
        X1 = Vert2(n).x
        z1 = Vert2(n).z
        X2 = InVert(n + 1).x
        z2 = InVert(n + 1).z
        If (z1 > p.z) Xor (z2 > p.z) Then
            If p.x > X1 + (X2 - X1) * (p.z - z1) / (z2 - z1) Then c = Not c
        End If
    Next n
    PointInSector = c
End Function

Friend Sub GetFaceUp(RetV3() As Map3Vertex, RetI() As Integer, RetVCnt As Long, RetICnt As Long, RetM As Long)
    Dim i As Long
    RetVCnt = vCnt
    RetICnt = iCnt
    RetM = MtrlUp
    ReDim RetI(iCnt - 1)
    For i = 0 To iCnt - 1
        RetI(i) = vInd(iCnt - 1 - i)
    Next i
    ReDim RetV3(vCnt - 1)
    For i = 0 To vCnt - 1
        RetV3(i).Pos.x = Vert2(i).x
        RetV3(i).Pos.y = yMax
        RetV3(i).Pos.z = Vert2(i).z
        RetV3(i).Norm.x = 0
        RetV3(i).Norm.y = 1
        RetV3(i).Norm.z = 0
        RetV3(i).tu = -(Vert2(i).x - xMin * (TexAlign(RetM))) * TexScale(RetM)
        RetV3(i).tv = -(Vert2(i).z - zMin * (TexAlign(RetM))) * TexScale(RetM)
    Next i
    CalcTBN RetV3(), vCnt
End Sub

Friend Sub GetFaceDown(RetV3() As Map3Vertex, RetI() As Integer, RetVCnt As Long, RetICnt As Long, RetM As Long)
    Dim i As Long
    RetVCnt = vCnt
    RetICnt = iCnt
    RetM = MtrlDown
    ReDim RetI(iCnt - 1)
    For i = 0 To iCnt - 1
        RetI(i) = vInd(i)
    Next i
    ReDim RetV3(vCnt - 1)
    For i = 0 To vCnt - 1
        RetV3(i).Pos.x = Vert2(i).x
        RetV3(i).Pos.y = yMin
        RetV3(i).Pos.z = Vert2(i).z
        RetV3(i).Norm.x = 0
        RetV3(i).Norm.y = -1
        RetV3(i).Norm.z = 0
        RetV3(i).tu = -(Vert2(i).x - xMin * (TexAlign(RetM))) * TexScale(RetM)
        RetV3(i).tv = (Vert2(i).z - zMin * (TexAlign(RetM))) * TexScale(RetM)
    Next i
    CalcTBN RetV3(), vCnt
End Sub

Friend Sub GetFace(ByVal Ind As Long, RetV3() As Map3Vertex, RetI() As Integer, RetVCnt As Long, RetICnt As Long, RetM As Long, RetT As Long)
    Static t As Single
    Dim x    As Single, z As Single, xz As Single
    Dim ii   As Long
    RetVCnt = 4
    RetICnt = 6
    RetM = Vert2(Ind).Mtrl
    RetT = Vert2(Ind).Target
    ReDim RetI(RetICnt - 1)
    RetI(0) = 0
    RetI(1) = 2
    RetI(2) = 1
    RetI(3) = 1
    RetI(4) = 2
    RetI(5) = 3
    ii = (Ind + 1) Mod vCnt
    x = Vert2(ii).x - Vert2(Ind).x
    z = Vert2(ii).z - Vert2(Ind).z
    xz = Sqr(x * x + z * z)
    ReDim RetV3(RetVCnt - 1)
    t = t * (1 - TexAlign(RetM))
    RetV3(0).Pos.x = Vert2(Ind).x
    RetV3(0).Pos.y = yMax
    RetV3(0).Pos.z = Vert2(Ind).z
    RetV3(0).Norm.x = z / xz
    RetV3(0).Norm.y = 0
    RetV3(0).Norm.z = -x / xz
    RetV3(0).tu = t * TexScale(RetM)
    RetV3(0).tv = yMax * (1 - TexAlign(RetM)) * TexScale(RetM)
    RetV3(1) = RetV3(0)
    RetV3(1).Pos.y = yMin
    RetV3(1).tv = (yMax - yMin) * TexScale(RetM) + RetV3(0).tv
    t = t - xz
    RetV3(2) = RetV3(0)
    RetV3(2).Pos.x = Vert2(ii).x
    RetV3(2).Pos.z = Vert2(ii).z
    RetV3(2).tu = t * TexScale(RetM)
    RetV3(3) = RetV3(2)
    RetV3(3).Pos.y = yMin
    RetV3(3).tv = RetV3(1).tv
    CalcTBN RetV3(), RetVCnt
End Sub

Public Function GetVCnt() As Long
    GetVCnt = vCnt
End Function

Private Sub CalcTBN(Vert3() As Map3Vertex, ByVal Size As Long)
    Dim i As Long
    For i = 0 To Size - 1
        Select Case Vert3(i).Norm.y
            Case Is < -0.9
                Vert3(i).Tang = Vec3(1, 0, 0)
                Vert3(i).Bnrm = Vec3(0, 0, -1)
            Case Is > 0.9
                Vert3(i).Tang = Vec3(1, 0, 0)
                Vert3(i).Bnrm = Vec3(0, 0, 1)
            Case Else
                Vert3(i).Bnrm = Vec3(0, 1, 0)
                Vec3Cross Vert3(i).Tang, Vert3(i).Norm, Vert3(i).Bnrm
        End Select
    Next i
End Sub
Private Sub Triangulator()
    Dim i   As Long, iL As Long, iR As Long, ii As Long
    Dim f() As Boolean, c As Long
    ReDim f(vCnt - 1)
    iCnt = 0
    ReDim vInd((vCnt - 2) * 3 - 1)
    c = vCnt
    Do
        While f(i)
            i = (i + 1) Mod vCnt
        Wend
        iR = (i + 1) Mod vCnt
        While f(iR)
            iR = (iR + 1) Mod vCnt
        Wend
        iL = (i + vCnt - 1) Mod vCnt
        While f(iL)
            iL = (iL + vCnt - 1) Mod vCnt
        Wend
        If c = 3 Then
            AddTri iL, i, iR, vInd(), iCnt
            Exit Do
        ElseIf CCW(Subt(Vert2(iL), Vert2(i)), Subt(Vert2(iR), Vert2(i))) < 0 Then
            ii = (iR + 1) Mod vCnt
            Do
                If ((ii + 1) Mod vCnt) = iL Then
                    If PointInTriangle(Vert2(ii), Vert2(i), Vert2(iR), Vert2(iL)) Then Exit Do
                    f(i) = True
                    AddTri iL, i, iR, vInd(), iCnt
                    c = c - 1
                    Exit Do
                ElseIf Intersect(Vert2(iL), Vert2(iR), Vert2(ii), Vert2((ii + 1) Mod vCnt)) Then
                    Exit Do
                Else
                    ii = (ii + 1) Mod vCnt
                End If
            Loop
        End If
        i = (i + 1) Mod vCnt
    Loop
End Sub
Private Sub AddTri(ByVal i1 As Long, ByVal i2 As Long, ByVal i3 As Long, vInd() As Integer, iCnt As Long)
    vInd(iCnt) = i1
    iCnt = iCnt + 1
    vInd(iCnt) = i2
    iCnt = iCnt + 1
    vInd(iCnt) = i3
    iCnt = iCnt + 1
End Sub
Private Function PointInTriangle(p As Map2Vertex, v1 As Map2Vertex, v2 As Map2Vertex, v3 As Map2Vertex) As Boolean
    If CCW(Subt(v2, v1), Subt(p, v1)) >= 0 And CCW(Subt(v3, v2), Subt(p, v2)) >= 0 And CCW(Subt(v1, v3), Subt(p, v3)) >= 0 Then PointInTriangle = True
End Function
Private Function Intersect(v1 As Map2Vertex, v2 As Map2Vertex, v3 As Map2Vertex, v4 As Map2Vertex) As Boolean
    Intersect = ((CCW(Subt(v2, v1), Subt(v3, v1)) <= 0) Xor (CCW(Subt(v2, v1), Subt(v4, v1)) <= 0)) And ((CCW(Subt(v4, v3), Subt(v2, v3)) <= 0) Xor (CCW(Subt(v4, v3), Subt(v1, v3)) <= 0))
End Function
Private Function Subt(v1 As Map2Vertex, v2 As Map2Vertex) As D3DVECTOR2
    Subt.x = v1.x - v2.x
    Subt.y = v1.z - v2.z
End Function
Private Function CCW(v1 As D3DVECTOR2, v2 As D3DVECTOR2) As Single
    CCW = v1.x * v2.y - v1.y * v2.x
End Function
