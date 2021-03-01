VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectSoundBuffer8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub dsb_Release Lib "dx9_vb" (ByVal pDSB As Long)
Private Declare Sub dsb_Play Lib "dx9_vb" (ByVal pDSB As Long, ByVal flags As Long)
Private Declare Sub dsb_Stop Lib "dx9_vb" (ByVal pDSB As Long)
Private Declare Function dsb_SetData Lib "dx9_vb" (ByVal pDSB As Long, ByVal Offset As Long, ByVal Size As Long, ByVal pData As Long) As Long
Private Declare Function dsb_GetData Lib "dx9_vb" (ByVal pDSB As Long, ByVal Offset As Long, ByVal Size As Long, ByVal pData As Long) As Long
Private Declare Sub dsb_SetCurrentPosition Lib "dx9_vb" (ByVal pDSB As Long, ByVal dwNewPosition As Long)
Private Declare Sub dsb_SetVolume Lib "dx9_vb" (ByVal pDSB As Long, ByVal Vol As Long)
Private Declare Sub dsb_SetPan Lib "dx9_vb" (ByVal pDSB As Long, ByVal Pan As Long)
Private Declare Sub dsb_SetFrequency Lib "dx9_vb" (ByVal pDSB As Long, ByVal Freq As Long)
Private mePtr As Long

Friend Sub SetVolume(ByVal Vol As Long)
    dsb_SetVolume Ptr, Vol
End Sub

Friend Sub SetPan(ByVal Pan As Long)
    dsb_SetPan Ptr, Pan
End Sub

Friend Sub SetFrequency(ByVal Freq As Long)
    dsb_SetFrequency Ptr, Freq
End Sub

Friend Sub SetCurrentPosition(ByVal dwNewPosition As Long)
    dsb_SetCurrentPosition Ptr, dwNewPosition
End Sub

Friend Function SetData(ByVal Offset As Long, ByVal Size As Long, ByVal pData As Long) As Boolean
    If dsb_SetData(Ptr, Offset, Size, pData) = DS_FALSE Then Exit Function
    SetData = True
End Function

Friend Function GetData(ByVal Offset As Long, ByVal Size As Long, ByVal pData As Long) As Boolean
    If dsb_GetData(Ptr, Offset, Size, pData) = DS_FALSE Then Exit Function
    GetData = True
End Function

Public Sub Play()
    dsb_Play Ptr, 0
End Sub

Public Sub PlayLooping()
    dsb_Play Ptr, 1
End Sub

Public Sub StopPlaying()
    dsb_Stop Ptr
End Sub

Private Sub Class_Terminate()
    If Ptr <> 0 Then
        dsb_Release Ptr
    End If
End Sub

Public Property Let Ptr(p As Long)
    If mePtr <> 0 Then dsb_Release mePtr
    mePtr = p
End Property
Public Property Get Ptr() As Long
    Ptr = mePtr
End Property
