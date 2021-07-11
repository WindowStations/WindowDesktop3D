VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Direct3DTexture9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub tex_Release Lib "dx9_vb" (ByVal pTex As Long)
Private Declare Sub tex_GetSurfaceLevel Lib "dx9_vb" (ByVal pTex As Long, ByVal Level As Long, ByRef pSurf As Long)
Private Declare Sub tex_GetLevelDesc Lib "dx9_vb" (ByVal pTex As Long, ByVal Level As Long, ByVal pDesc As Long)
Private Declare Sub tex_GenerateMipSubLevels Lib "dx9_vb" (ByVal pTex As Long)
Private mePtr As Long
Friend Sub GenerateMipSubLevels()
   tex_GenerateMipSubLevels Ptr
End Sub
Friend Function GetLevelDesc(ByVal Level As Long) As D3DSURFACE_DESC
   Dim sd As D3DSURFACE_DESC
   If Ptr <> 0 Then
      tex_GetLevelDesc Ptr, Level, VarPtr(sd)
      GetLevelDesc = sd
   End If
End Function
Friend Function GetSurfaceLevel(ByVal Level As Long) As Direct3DSurface9
   Dim pSurf As Long
   If Ptr <> 0 Then
      tex_GetSurfaceLevel Ptr, Level, pSurf
      If pSurf <> 0 Then
         Set GetSurfaceLevel = New Direct3DSurface9
         GetSurfaceLevel.Ptr = pSurf
      End If
   End If
End Function
Private Sub Class_Terminate()
   If Ptr <> 0 Then
      tex_Release Ptr
   End If
End Sub
Public Property Let Ptr(p As Long)
   If mePtr <> 0 Then tex_Release mePtr
   mePtr = p
End Property
Public Property Get Ptr() As Long
   Ptr = mePtr
End Property
