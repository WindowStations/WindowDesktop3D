VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Direct3D9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub d3d_CreateDevice Lib "dx9_vb" (ByVal pD3d As Long, ByVal hWnd As Long, ByVal flags As CONST_D3DCREATEFLAGS, ByRef PP As D3DPRESENT_PARAMETERS, ByRef pDev As Long)
Private Declare Sub d3d_Release Lib "dx9_vb" (ByVal pD3d As Long)
Private Declare Sub d3d_GetDeviceCaps Lib "dx9_vb" (ByVal pD3d As Long, ByVal Adapter As Long, ByVal DeviceType As D3DDEVTYPE, ByRef Caps As D3DCAPS9)
Private mePtr As Long
Friend Sub GetDeviceCaps(ByVal Adapter As Long, ByVal DeviceType As D3DDEVTYPE, ByRef Caps As D3DCAPS9)
   d3d_GetDeviceCaps Ptr, Adapter, DeviceType, Caps
End Sub
Friend Function CreateDevice(ByVal hWnd As Long, ByVal flags As CONST_D3DCREATEFLAGS, ByRef d3dpp As D3DPRESENT_PARAMETERS) As Direct3DDevice9
   Dim pDev As Long
   d3d_CreateDevice Ptr, hWnd, flags, d3dpp, pDev
   If pDev <> 0 Then
      Set CreateDevice = New Direct3DDevice9
      CreateDevice.Ptr = pDev
   End If
End Function
Public Property Let Ptr(p As Long)
   If mePtr <> 0 Then d3d_Release mePtr
   mePtr = p
End Property
Public Property Get Ptr() As Long
   Ptr = mePtr
End Property
Private Sub Class_Terminate()
   If Ptr <> 0 Then
      d3d_Release Ptr
   End If
End Sub
