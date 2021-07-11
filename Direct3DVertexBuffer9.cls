VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Direct3DVertexBuffer9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub vb_Release Lib "dx9_vb" (ByVal pVB As Long)
Private Declare Function vb_SetData Lib "dx9_vb" (ByVal pVB As Long, ByVal Offset As Long, ByVal Size As Long, ByVal pData As Long, ByVal flags As Long) As Long
Private Declare Function vb_GetData Lib "dx9_vb" (ByVal pVB As Long, ByVal Offset As Long, ByVal Size As Long, ByVal pData As Long, ByVal flags As Long) As Long
Private mePtr As Long
Friend Function SetData(ByVal Offset As Long, ByVal Size As Long, ByVal pData As Long, ByVal flags As Long) As Boolean
   If vb_SetData(Ptr, Offset, Size, pData, flags) = D3D_FALSE Then Exit Function
   SetData = True
End Function
Friend Function GetData(ByVal Offset As Long, ByVal Size As Long, ByVal pData As Long, ByVal flags As Long) As Boolean
   If vb_GetData(Ptr, Offset, Size, pData, flags) = D3D_FALSE Then Exit Function
   GetData = True
End Function
Private Sub Class_Terminate()
   If Ptr <> 0 Then
      vb_Release Ptr
   End If
End Sub
Public Property Let Ptr(p As Long)
   If mePtr <> 0 Then vb_Release mePtr
   mePtr = p
End Property
Public Property Get Ptr() As Long
   Ptr = mePtr
End Property
