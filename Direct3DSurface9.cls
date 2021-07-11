VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Direct3DSurface9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub surf_Release Lib "dx9_vb" (ByVal pSurf As Long)
Private Declare Function surf_SetData Lib "dx9_vb" (ByVal pSurf As Long, ByVal pRect As Long, ByVal Size As Long, ByVal pData As Long) As Long
Private Declare Function surf_GetData Lib "dx9_vb" (ByVal pSurf As Long, ByVal pRect As Long, ByVal Size As Long, ByVal pData As Long) As Long
Private Declare Function surf_GetPixel Lib "dx9_vb" (ByVal pSurf As Long, ByVal x As Long, ByVal y As Long, ByVal Size As Long, ByVal pData As Long) As Long
Private Declare Sub surf_GetDesc Lib "dx9_vb" (ByVal pSurf As Long, ByVal pDesc As Long)
Private mePtr As Long
Friend Function GetDesc() As D3DSURFACE_DESC
   Dim sd As D3DSURFACE_DESC
   If Ptr <> 0 Then
      surf_GetDesc Ptr, VarPtr(sd)
      GetDesc = sd
   End If
End Function
Friend Function SetData(ByVal pRect As Long, ByVal Size As Long, ByVal pData As Long) As Boolean
   If surf_SetData(Ptr, pRect, Size, pData) = D3D_TRUE Then SetData = True
End Function
Friend Function GetData(ByVal pRect As Long, ByVal Size As Long, ByVal pData As Long) As Boolean
   If surf_GetData(Ptr, pRect, Size, pData) = D3D_TRUE Then GetData = True
End Function
Friend Function GetPixel(ByVal x As Long, ByVal y As Long, ByVal Size As Long, ByVal pData As Long) As Boolean
   If surf_GetPixel(Ptr, x, y, Size, pData) = D3D_TRUE Then GetPixel = True
End Function
Private Sub Class_Terminate()
   If Ptr <> 0 Then
      surf_Release Ptr
   End If
End Sub
Public Property Let Ptr(p As Long)
   If mePtr <> 0 Then surf_Release mePtr
   mePtr = p
End Property
Public Property Get Ptr() As Long
   Ptr = mePtr
End Property
