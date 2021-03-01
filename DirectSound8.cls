VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectSound8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub ds_Release Lib "dx9_vb" (ByVal pDS As Long)
Private Declare Sub ds_CreateSoundBuffer Lib "dx9_vb" (ByVal pDS As Long, ByVal nChannels As Long, ByVal BitsPerSample As Long, ByVal Freq As Long, ByVal Size As Long, ByVal flags As CONST_DSBCAPS, ByRef pDSB As Long)
Private Declare Sub ds_DuplicateSoundBuffer Lib "dx9_vb" (ByVal pDS As Long, ByVal pDSBIn As Long, ByRef pDSBOut As Long)
Private Declare Sub ds_SetCooperativeLevel Lib "dx9_vb" (ByVal pDS As Long, ByVal hWnd As Long, ByVal dwLevel As CONST_DSSCL)
Private mePtr As Long

Friend Sub SetCooperativeLevel(ByVal hWnd As Long, ByVal dwLevel As CONST_DSSCL)
    ds_SetCooperativeLevel Ptr, hWnd, dwLevel
End Sub

Friend Function DuplicateSoundBuffer(dsb As DirectSoundBuffer8) As DirectSoundBuffer8
    Dim pDSB As Long
    ds_DuplicateSoundBuffer Ptr, dsb.Ptr, pDSB
    If pDSB <> 0 Then
        Set DuplicateSoundBuffer = New DirectSoundBuffer8
        DuplicateSoundBuffer.Ptr = pDSB
    End If
End Function

Friend Function CreateSoundBuffer(ByVal nChannels As Long, ByVal BitsPerSample As Long, ByVal Freq As Long, ByVal Size As Long, ByVal flags As CONST_DSBCAPS) As DirectSoundBuffer8
    Dim pDSB As Long
    ds_CreateSoundBuffer Ptr, nChannels, BitsPerSample, Freq, Size, flags, pDSB
    If pDSB <> 0 Then
        Set CreateSoundBuffer = New DirectSoundBuffer8
        CreateSoundBuffer.Ptr = pDSB
    End If
End Function

Private Sub Class_Terminate()
    If Ptr <> 0 Then
        ds_Release Ptr
    End If
End Sub

Public Property Let Ptr(p As Long)
    If mePtr <> 0 Then ds_Release mePtr
    mePtr = p
End Property
Public Property Get Ptr() As Long
    Ptr = mePtr
End Property
