VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProjectile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event OnStrike(ByVal iProjectile As Long, ByVal jButton As Long, ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal r As Single)
Public Event OnMiss(ByVal iProjectile As Long, ByVal jButton As Long, ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal r As Single)
Public Sub Miss(ByVal iProjectile As Long, ByVal jButton As Long, ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal r As Single)
   On Error Resume Next
   RaiseEvent OnMiss(iProjectile, jButton, x, y, z, r)
End Sub
Public Sub Strike(ByVal iProjectile As Long, ByVal jButton As Long, ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal r As Single)
   On Error Resume Next
   RaiseEvent OnStrike(iProjectile, jButton, x, y, z, r)
End Sub
Friend Function add(p As D3DVECTOR, s As D3DVECTOR) As D3DVECTOR
   On Error Resume Next
   projectileselecting = True
   Projectiles.Pos = p
   Projectiles.Speed = s
   Projectiles.Timer = 1000
   'CloudT(0).Spd = s
   Dim v  As D3DVECTOR
   Dim v2 As D3DVECTOR
   Dim j  As Long
   Dim i  As Long
   Dim b As Boolean
   For i = 1 To 1000
      Vec3Add v, Projectiles.Pos, Projectiles.Speed 'i'calculate the position based on speed
      Projectiles.Pos = v 'set new position
      '
      ' Vec3Scale v, Projectiles.Speed, -1
      '
      '
      Vec3Scale v, Projectiles.Speed, 0.07 '?
      '
      '
      Projectiles.Timer = Projectiles.Timer - 1 'decrement life span of projectile
      '        For j = 0 To ButCnt - 1 'loop through buttons preloaded with map load
      '            v2 = v 'get reset the freshest vector to calculate from
      '            Vec3Subtract v2, But(j).Pos, Projectiles.Pos 'calculate distance between button and projectile
      '            If Vec3LengthSq(v2) < 5 Then 'if distance is less than specified strike distance
      '                Projectiles.Timer = 0 'projectile hit a target, end life span
      '                Strike 0, j, Projectiles.Pos.x, Projectiles.Pos.y, Projectiles.Pos.z, Rand 'signal event outwards to main game logic
      '                Exit For
      '            End If
      '        Next
      For j = 0 To ThingCnt - 1 'loop through any inserted objects in the same way as buttons above
         v2 = v
         Vec3Subtract v2, Thing(j).Pos, Projectiles.Pos
         If Vec3LengthSq(v2) < 3 Then
            b = True
            Projectiles.Timer = 0
            Strike 0, j, Projectiles.Pos.x, Projectiles.Pos.y, Projectiles.Pos.z, Rand
            add = Projectiles.Pos
            Exit For
         End If
      Next
      If b = True Then Exit For
   Next
   If b = False Then 'if selection was not an object
      Miss 0, j, Projectiles.Pos.x, Projectiles.Pos.y, Projectiles.Pos.z, Rand
   End If
End Function
Friend Sub Remove(ByVal Ind As Long)
   On Error Resume Next
   projectileselecting = False
   '    frmMain.targetClass.Add Projectiles.Pos, 10 'ind
End Sub
'Friend Sub Add(p As D3DVECTOR, s As D3DVECTOR)
'    On Error Resume Next
'    projectileselecting = True
'    Projectiles.Pos = p
'    Projectiles.Speed = s
'    Projectiles.Timer = 1000
'End Sub
'
'Friend Sub Tick()
'    On Error Resume Next
'    Dim v  As D3DVECTOR
'    Dim v2 As D3DVECTOR
'    Dim j  As Long
'
'
'    Vec3Add v, Projectiles.Pos, Projectiles.Speed 'i'calculate the position based on speed
''    Vec3Add v, v, Projectiles.Speed
''    Vec3Add v, v, Projectiles.Speed
''    Vec3Add v, v, Projectiles.Speed
'
'
'    Projectiles.Pos = v 'set new position
'    '
'    ' Vec3Scale v, Projectiles.Speed, -1
'    '
'    Vec3Scale v, Projectiles.Speed, 0.07 '?
'    '
'    '
'    Projectiles.Timer = Projectiles.Timer - 1 'decrement life span of projectile
'    For j = 0 To ButCnt - 1 'loop through buttons preloaded with map load
'        v2 = v 'get reset the freshest vector to calculate from
'        Vec3Subtract v2, But(j).Pos, Projectiles.Pos 'calculate distance between button and projectile
'        If Vec3LengthSq(v2) < 10 Then 'if distance is less than specified strike distance
'            Projectiles.Timer = 0 'projectile hit a target, end life span
'            Strike 0, j, Projectiles.Pos.x, Projectiles.Pos.y, Projectiles.Pos.z, Rand 'signal event outwards to main game logic
'        End If
'    Next
'    For j = 0 To ThingCnt - 1 'loop through any inserted objects in the same way as buttons above
'        v2 = v
'        Vec3Subtract v2, Thing(j).Pos, Projectiles.Pos
'        If Vec3LengthSq(v2) < 10 Then
'            Projectiles.Timer = 0
'            Strike 0, j, Projectiles.Pos.x, Projectiles.Pos.y, Projectiles.Pos.z, Rand
'        End If
'    Next
'End Sub
'Friend Sub Remove(ByVal Ind As Long)
'    On Error Resume Next
'    projectileselecting = False
''    frmMain.targetClass.Add Projectiles.Pos, 10 'ind
'End Sub
