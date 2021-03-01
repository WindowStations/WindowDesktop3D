VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectSound3DBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim dsb   As DirectSoundBuffer8
Dim meVol As Single
Dim meRad As Single
Dim kVol  As Single
Dim kPan  As Single

Friend Sub Init(DSB8 As DirectSoundBuffer8, ByVal Radius As Single)
    Set dsb = DSB8
    meRad = Radius
End Sub

Friend Sub Update(Pos As D3DVECTOR, vEye As D3DVECTOR, vDir As D3DVECTOR, vUp As D3DVECTOR)
    Dim vR As D3DVECTOR, vD As D3DVECTOR
    Vec3Subtract vD, Pos, vEye
    kVol = Sqr(meRad / (Vec3LengthSq(vD) + meRad))
    Vec3Cross vR, vUp, vDir
    Vec3Normalize vD, vD
    kPan = Vec3Dot(vR, vD) * (1 - kVol)
    SetParm
End Sub

Private Sub SetParm()
    Dim v As Long
    Dim s As Single
    s = Log(meVol * kVol * 9999 + 1)
    v = s * 1085 - 10000
    dsb.SetVolume v
    dsb.SetPan kPan * 1000
End Sub

Public Sub Play()
    dsb.SetCurrentPosition 0
    dsb.Play
End Sub

Public Sub PlayLoop()
    dsb.SetCurrentPosition 0
    dsb.PlayLooping
End Sub

Public Sub StopPlaying()
    dsb.StopPlaying
End Sub

Public Sub SetFreq(ByVal Freq As Long)
    dsb.SetFrequency Freq
End Sub

Public Sub SetVol(ByVal Vol As Single)
    Select Case Vol
        Case Is < 0
            meVol = 0
        Case Is > 1
            meVol = 1
        Case Else
            meVol = Vol
    End Select
    SetParm
End Sub

Private Sub Class_Initialize()
    meVol = 1
    kVol = 1
    kPan = 0
End Sub

Private Sub Class_Terminate()
    dsb.StopPlaying
    Set dsb = Nothing
End Sub
