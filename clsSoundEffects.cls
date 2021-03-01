VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSoundEffects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private DS8          As DirectSound8
Private DSBDoor(3)   As DirectSoundBuffer8
Private DS3DDoor(3)  As DirectSound3DBuffer
Private DoorTimer(3) As Long
Private DoorInd(3)   As Long
Private DSBDoorStop  As DirectSoundBuffer8
Private DS3DDoorStop As DirectSound3DBuffer
Private DSBBang(3)   As DirectSoundBuffer8
Private DS3DBang(3)  As DirectSound3DBuffer
Private BangInd      As Long
Private DSBBomb(1)   As DirectSoundBuffer8
Private DS3DBomb(1)  As DirectSound3DBuffer
Private BombInd      As Long
'Private DSBWind      As DirectSoundBuffer8
Private DSBButOn     As DirectSoundBuffer8
Private DSBButFail   As DirectSoundBuffer8
Private DSBw1        As DirectSoundBuffer8
Private DSBw3        As DirectSoundBuffer8
Private DSBwR1       As DirectSoundBuffer8
Private DSBwR2       As DirectSoundBuffer8
Private DSBwR3       As DirectSoundBuffer8
Private DSBMed       As DirectSoundBuffer8
Private DSBStep      As DirectSoundBuffer8
Private DSBDamage    As DirectSoundBuffer8
Private StepTimer    As Long
Private BombTimer    As Long
Friend Sub Initialize(ByVal hWnd As Long)
    On Error Resume Next
    Dim ar() As Integer
    Dim nf   As Integer
    Dim sz   As Long
    Dim i    As Long
    Set DS8 = CreateDirectSound
    DS8.SetCooperativeLevel hWnd, DSSCL_NORMAL
    nf = FreeFile
    Open App.Path & "\3DEngine\Sound\Door.wav" For Binary As #nf
    Get #nf, &H29, sz
    ReDim ar(sz \ 2 - 1)
    Get #nf, &H2D, ar()
    Close #nf
    Set DSBDoor(0) = DS8.CreateSoundBuffer(1, 16, 11025, sz, DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLFREQUENCY)
    DSBDoor(0).SetData 0, sz, VarPtr(ar(0))
    For i = 1 To 3
        Set DSBDoor(i) = DS8.DuplicateSoundBuffer(DSBDoor(0))
    Next i
    For i = 0 To 3
        Set DS3DDoor(i) = New DirectSound3DBuffer
        DS3DDoor(i).Init DSBDoor(i), 50
        DS3DDoor(i).SetFreq 8000 + i * 200
        DS3DDoor(i).SetVol 0.55
    Next i
    '
    nf = FreeFile
    Open App.Path & "\3DEngine\Sound\Bang.wav" For Binary As #nf
    Get #nf, &H29, sz
    ReDim ar(sz \ 2 - 1)
    Get #nf, &H2D, ar()
    Close #nf
    Set DSBBang(0) = DS8.CreateSoundBuffer(1, 16, 44100, sz, DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLFREQUENCY)
    DSBBang(0).SetData 0, sz, VarPtr(ar(0))
    For i = 1 To 3
        Set DSBBang(i) = DS8.DuplicateSoundBuffer(DSBBang(0))
    Next i
    For i = 0 To 3
        Set DS3DBang(i) = New DirectSound3DBuffer
        DS3DBang(i).Init DSBBang(i), 150
    Next i
    '
    '    nf = FreeFile
    '    Open App.Path & "\3DEngine\Sound\Bomb.wav" For Binary As #nf
    '    Get #nf, &H29, sz
    '    ReDim ar(sz \ 2 - 1)
    '    Get #nf, &H2D, ar()
    '    Close #nf
    '    Set DSBBomb(0) = DS8.CreateSoundBuffer(1, 16, 22050, sz, DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLFREQUENCY)
    '    DSBBomb(0).SetData 0, sz, VarPtr(ar(0))
    '    Set DSBBomb(1) = DS8.DuplicateSoundBuffer(DSBBomb(0))
    '    For i = 0 To 1
    '        Set DS3DBomb(i) = New DirectSound3DBuffer
    '        DS3DBomb(i).Init DSBBomb(i), 15
    '        DS3DBomb(i).SetVol 1
    '    Next i
    '    nf = FreeFile
    '    Open App.Path & "\3DEngine\Sound\DoorStop.wav" For Binary As #nf
    '    Get #nf, &H29, sz
    '    ReDim ar(sz \ 2 - 1)
    '    Get #nf, &H2D, ar()
    '    Close #nf
    '    Set DSBDoorStop = DS8.CreateSoundBuffer(1, 16, 44100, sz, DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLFREQUENCY)
    '    DSBDoorStop.SetData 0, sz, VarPtr(ar(0))
    '    Set DS3DDoorStop = New DirectSound3DBuffer
    '    DS3DDoorStop.Init DSBDoorStop, 50
    '    DS3DDoorStop.SetFreq 44100
    '    DS3DDoorStop.SetVol 0.6
    nf = FreeFile
    Open App.Path & "\3DEngine\Sound\ButtonOn.wav" For Binary As #nf
    Get #nf, &H29, sz
    ReDim ar(sz \ 2 - 1)
    Get #nf, &H2D, ar()
    Close #nf
    Set DSBButOn = DS8.CreateSoundBuffer(1, 16, 22050, sz, DSBCAPS_CTRLVOLUME)
    DSBButOn.SetData 0, sz, VarPtr(ar(0))
    '
    '    nf = FreeFile
    '    Open App.Path & "\3DEngine\Sound\CheckT1.wav" For Binary As #nf
    '    Get #nf, &H29, sz
    '    ReDim ar(sz \ 2 - 1)
    '    Get #nf, &H2D, ar()
    '    Close #nf
    '    Set DSBWind = DS8.CreateSoundBuffer(2, 16, 22050, sz, DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLFREQUENCY)
    '    'Set DSBWind = DS8.CreateSoundBuffer(1, 16, 44100, sz, DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLFREQUENCY)
    '    'Set DSBWind = DS8.CreateSoundBuffer(1, 16, 22050, sz, DSBCAPS_CTRLVOLUME)
    '    DSBWind.SetData 0, sz, VarPtr(ar(0))
    '
    nf = FreeFile
    Open App.Path & "\3DEngine\Sound\ButtonFail.wav" For Binary As #nf
    Get #nf, &H29, sz
    ReDim ar(sz \ 2 - 1)
    Get #nf, &H2D, ar()
    Close #nf
    Set DSBButFail = DS8.CreateSoundBuffer(1, 16, 22050, sz, DSBCAPS_CTRLVOLUME)
    DSBButFail.SetData 0, sz, VarPtr(ar(0))
    '
    nf = FreeFile
    Open App.Path & "\3DEngine\Sound\Step.wav" For Binary As #nf
    Get #nf, &H29, sz
    ReDim ar(sz \ 2 - 1)
    Get #nf, &H2D, ar()
    Close #nf
    Set DSBStep = DS8.CreateSoundBuffer(1, 16, 44100, sz, DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLFREQUENCY)
    DSBStep.SetData 0, sz, VarPtr(ar(0))
    '
    nf = FreeFile
    Open App.Path & "\3DEngine\Sound\w1.wav" For Binary As #nf
    Get #nf, &H29, sz
    ReDim ar(sz \ 2 - 1)
    Get #nf, &H2D, ar()
    Close #nf
    Set DSBw1 = DS8.CreateSoundBuffer(1, 16, 22050, sz, DSBCAPS_CTRLVOLUME)
    DSBw1.SetData 0, sz, VarPtr(ar(0))
    DSBw1.SetVolume -1100
    '
    nf = FreeFile
    Open App.Path & "\3DEngine\Sound\w3.wav" For Binary As #nf
    Get #nf, &H29, sz
    ReDim ar(sz \ 2 - 1)
    Get #nf, &H2D, ar()
    Close #nf
    Set DSBw3 = DS8.CreateSoundBuffer(1, 16, 14000, sz, DSBCAPS_CTRLVOLUME)
    DSBw3.SetData 0, sz, VarPtr(ar(0))
    DSBw3.SetVolume 0
    '
    nf = FreeFile
    Open App.Path & "\3DEngine\Sound\wR1.wav" For Binary As #nf
    Get #nf, &H29, sz
    ReDim ar(sz \ 2 - 1)
    Get #nf, &H2D, ar()
    Close #nf
    Set DSBwR1 = DS8.CreateSoundBuffer(1, 16, 22050, sz, DSBCAPS_CTRLVOLUME)
    DSBwR1.SetData 0, sz, VarPtr(ar(0))
    '
    nf = FreeFile
    Open App.Path & "\3DEngine\Sound\wR2.wav" For Binary As #nf
    Get #nf, &H29, sz
    ReDim ar(sz \ 2 - 1)
    Get #nf, &H2D, ar()
    Close #nf
    Set DSBwR3 = DS8.CreateSoundBuffer(1, 16, 16000, sz, DSBCAPS_CTRLVOLUME)
    DSBwR3.SetData 0, sz, VarPtr(ar(0))
    '
    nf = FreeFile
    Open App.Path & "\3DEngine\Sound\Med.wav" For Binary As #nf
    Get #nf, &H29, sz
    ReDim ar(sz \ 2 - 1)
    Get #nf, &H2D, ar()
    Close #nf
    Set DSBMed = DS8.CreateSoundBuffer(1, 16, 22050, sz, DSBCAPS_CTRLVOLUME)
    DSBMed.SetData 0, sz, VarPtr(ar(0))
    DSBMed.SetVolume -1000
    '
    nf = FreeFile
    Open App.Path & "\3DEngine\Sound\Damage.wav" For Binary As #nf
    Get #nf, &H29, sz
    ReDim ar(sz \ 2 - 1)
    Get #nf, &H2D, ar()
    Close #nf
    Set DSBDamage = DS8.CreateSoundBuffer(1, 16, 36000, sz, DSBCAPS_CTRLVOLUME)
    DSBDamage.SetData 0, sz, VarPtr(ar(0))
    DSBDamage.SetVolume -700
End Sub
Friend Sub Tick()
    On Error Resume Next
    Dim i As Long
    For i = 0 To 3
        If DoorTimer(i) = 20 Then
            DoorTimer(i) = DoorTimer(i) - 1
            DS3DDoor(i).Update Door(DoorInd(i)).Pos, PlPos, PlDir, Vec3(0, 1, 0)
            DS3DDoorStop.Update Door(DoorInd(i)).Pos, PlPos, PlDir, Vec3(0, 1, 0)
            DS3DDoorStop.Play
        ElseIf DoorTimer(i) > 1 Then
            DoorTimer(i) = DoorTimer(i) - 1
            DS3DDoor(i).Update Door(DoorInd(i)).Pos, PlPos, PlDir, Vec3(0, 1, 0)
        ElseIf DoorTimer(i) = 1 Then
            DS3DDoor(i).StopPlaying
            DoorTimer(i) = 0
        End If
    Next i
    If StepTimer > 0 Then StepTimer = StepTimer - 1
    If BombTimer > 0 Then BombTimer = BombTimer - 1
End Sub
Friend Sub SelectItem(Pos As D3DVECTOR, ByVal Vol As Single, ByVal Freq As Long)
    On Error Resume Next
    If soundxMute = 1 Then Exit Sub
    DS3DBang(BangInd).SetVol Vol
    DS3DBang(BangInd).SetFreq Freq
    DS3DBang(BangInd).Update Pos, PlPos, PlDir, Vec3(0, 1, 0)
    DS3DBang(BangInd).Play
    BangInd = (BangInd + 1) Mod 4
End Sub
'Friend Sub Bomb(Pos As D3DVECTOR)
'    On Error Resume Next
'    Dim v As D3DVECTOR
'    If BombTimer > 0 Then Exit Sub
'    Vec3Subtract v, Pos, PlPos
'    If Vec3LengthSq(v) > 10000 Then Exit Sub
'    DS3DBomb(BombInd).SetFreq 11000 - Rand * 200
'    DS3DBomb(BombInd).Update Pos, PlPos, PlDir, Vec3(0, 1, 0)
'    DS3DBomb(BombInd).Play
'    BombInd = 1 - BombInd
'    BombTimer = 20
'End Sub
'Friend Sub DoorOperate(ByVal t As Long, Ind As Long)
'    On Error Resume Next
'    Dim i As Long
'    For i = 0 To 2
'        If DoorTimer(i) = 0 Then Exit For
'    Next i
'    DoorTimer(i) = t
'    DoorInd(i) = Ind
'    DS3DDoor(i).Update Door(DoorInd(i)).Pos, PlPos, PlDir, Vec3(0, 1, 0)
'    DS3DDoor(i).PlayLoop
'End Sub
Friend Sub ButtonOn()
    On Error Resume Next
    If soundxMute = 1 Then Exit Sub
    DSBButOn.StopPlaying
    DSBButOn.SetCurrentPosition 0
    DSBButOn.Play
End Sub
'Friend Sub Windrunning()
'    On Error Resume Next
'    DSBWind.SetVolume -200
'
'    DSBWind.StopPlaying
'    DSBWind.SetCurrentPosition 0
'    DSBWind.Play
'End Sub
Friend Sub ButtonFail()
    On Error Resume Next
    If soundxMute = 1 Then Exit Sub
    DSBButFail.StopPlaying
    DSBButFail.SetCurrentPosition 0
    DSBButFail.Play
End Sub
Friend Sub Step(ByVal Vol As Single)
    On Error Resume Next
    If soundxMute = 1 Then Exit Sub
    ' If StepTimer > 0 Then Exit Sub
    DSBStep.StopPlaying
    DSBStep.SetVolume (Vol - 1) * 2500
    DSBStep.SetCurrentPosition 0
    DSBStep.SetFrequency 44100 - Rand * 5000
    DSBStep.Play
    StepTimer = 400
End Sub
Friend Sub Weapon1()
    On Error Resume Next
    If soundxMute = 1 Then Exit Sub
    DSBw1.StopPlaying
    DSBw1.SetCurrentPosition 0
    DSBw1.Play
End Sub
Friend Sub Weapon3()
    On Error Resume Next
    If soundxMute = 1 Then Exit Sub
    DSBw3.StopPlaying
    DSBw3.SetCurrentPosition 0
    DSBw3.Play
End Sub
Friend Sub WeaponR1()
    On Error Resume Next
    If soundxMute = 1 Then Exit Sub
    DSBwR1.StopPlaying
    DSBwR1.SetCurrentPosition 0
    DSBwR1.Play
End Sub
Friend Sub WeaponR2()
    On Error Resume Next
    If soundxMute = 1 Then Exit Sub
    DSBwR2.StopPlaying
    DSBwR2.SetCurrentPosition 0
    DSBwR2.Play
End Sub
Friend Sub WeaponR3()
    On Error Resume Next
    If soundxMute = 1 Then Exit Sub
    DSBwR3.StopPlaying
    DSBwR3.SetCurrentPosition 0
    DSBwR3.Play
End Sub
Friend Sub Medkit()
    On Error Resume Next
    If soundxMute = 1 Then Exit Sub
    DSBMed.StopPlaying
    DSBMed.SetCurrentPosition 0
    DSBMed.Play
End Sub
Friend Sub Damage()
    On Error Resume Next
    If soundxMute = 1 Then Exit Sub
    DSBDamage.StopPlaying
    DSBDamage.SetCurrentPosition 0
    DSBDamage.Play
End Sub
Friend Sub StopPlay()
    On Error Resume Next
    Dim i As Long
    For i = 0 To 3
        DS3DDoor(i).StopPlaying
        DoorTimer(i) = 0
    Next i
End Sub
Friend Sub Terminate()
    On Error Resume Next
    Dim i As Long
    For i = 0 To 3
        DS3DDoor(i).StopPlaying
        Set DS3DDoor(i) = Nothing
        Set DSBDoor(i) = Nothing
    Next i
    DS3DDoorStop.StopPlaying
    Set DS3DDoorStop = Nothing
    Set DSBDoorStop = Nothing
    For i = 0 To 3
        DS3DBang(i).StopPlaying
        Set DS3DBang(i) = Nothing
        Set DSBBang(i) = Nothing
    Next i
    For i = 0 To 1
        DS3DBomb(i).StopPlaying
        Set DS3DBomb(i) = Nothing
        Set DSBBomb(i) = Nothing
    Next i
    DSBButOn.StopPlaying
    Set DSBButOn = Nothing
    DSBButFail.StopPlaying
    Set DSBButFail = Nothing
    DSBw1.StopPlaying
    Set DSBw1 = Nothing
    DSBw3.StopPlaying
    Set DSBw3 = Nothing
    DSBwR3.StopPlaying
    Set DSBwR3 = Nothing
    'DSBwR2.StopPlaying
    Set DSBwR2 = Nothing
    DSBwR1.StopPlaying
    Set DSBwR1 = Nothing
    DSBMed.StopPlaying
    Set DSBMed = Nothing
    DSBStep.StopPlaying
    Set DSBStep = Nothing
    DSBDamage.StopPlaying
    Set DSBDamage = Nothing
    Set DS8 = Nothing
End Sub
