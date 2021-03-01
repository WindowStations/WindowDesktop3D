VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Direct3DIndexBuffer9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub ib_Release Lib "dx9_vb" (ByVal pIB As Long)
Private Declare Function ib_SetData Lib "dx9_vb" (ByVal pIB As Long, ByVal Offset As Long, ByVal Size As Long, ByVal pData As Long, ByVal flags As Long) As Long
Private Declare Function ib_GetData Lib "dx9_vb" (ByVal pIB As Long, ByVal Offset As Long, ByVal Size As Long, ByVal pData As Long, ByVal flags As Long) As Long
Private mePtr As Long

Friend Function SetData(ByVal Offset As Long, ByVal Size As Long, ByVal pData As Long, ByVal flags As Long) As Boolean
    If ib_SetData(Ptr, Offset, Size, pData, flags) = D3D_FALSE Then Exit Function
    SetData = True
End Function

Friend Function GetData(ByVal Offset As Long, ByVal Size As Long, ByVal pData As Long, ByVal flags As Long) As Boolean
    If ib_GetData(Ptr, Offset, Size, pData, flags) = D3D_FALSE Then Exit Function
    GetData = True
End Function

Private Sub Class_Terminate()
    If Ptr <> 0 Then
        ib_Release Ptr
    End If
End Sub

Public Property Let Ptr(p As Long)
    If mePtr <> 0 Then ib_Release mePtr
    mePtr = p
End Property
Public Property Get Ptr() As Long
    Ptr = mePtr
End Property
