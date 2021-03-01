VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Direct3DVolumeTexture9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub vtex_Release Lib "dx9_vb" (ByVal pTex As Long)
Private Declare Sub vtex_GetVolumeLevel Lib "dx9_vb" (ByVal pTex As Long, ByVal Level As Long, ByRef pVol As Long)
Private Declare Sub vtex_GenerateMipSubLevels Lib "dx9_vb" (ByVal pTex As Long)
Private mePtr As Long
 
Friend Sub GenerateMipSubLevels()
    vtex_GenerateMipSubLevels Ptr
End Sub
 
Friend Function GetVolumeLevel(ByVal Level As Long) As Direct3DVolume9
    Dim pVol As Long
    If Ptr <> 0 Then
        vtex_GetVolumeLevel Ptr, Level, pVol
        If pVol <> 0 Then
            Set GetVolumeLevel = New Direct3DVolume9
            GetVolumeLevel.Ptr = pVol
        End If
    End If
End Function

Private Sub Class_Terminate()
    If Ptr <> 0 Then
        vtex_Release Ptr
    End If
End Sub

Public Property Let Ptr(p As Long)
    If mePtr <> 0 Then vtex_Release mePtr
    mePtr = p
End Property
Public Property Get Ptr() As Long
    Ptr = mePtr
End Property
