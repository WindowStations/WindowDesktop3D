VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAVIToPictureBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const WS_CHILD          As Long = &H40000000
Private Const WS_VISIBLE        As Long = &H10000000
Private Const WM_USER           As Long = &H400
Private Const ICC_ANIMATE_CLASS As Long = &H80
Private Const ANIMATE_CLASSA    As String = "SysAnimate32"
Private Const ACS_CENTER        As Long = &H1
Private Const ACS_TRANSPARENT   As Long = &H2
Private Const ACS_AUTOPLAY      As Long = &H4
Private Const ACS_TIMER         As Long = &H8
Private Const ACM_OPEN          As Long = (WM_USER + 100)
Private Const ACM_PLAY          As Long = (WM_USER + 101)
Private Const ACM_STOP          As Long = (WM_USER + 102)
Private Type tagInitCommonControlsEx
   dwSize As Long
   dwICC As Long
End Type
Private Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type
Private Declare Function apiLoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function apiFreeLibrary Lib "kernel32" Alias "FreeLibrary" (ByVal hLibModule As Long) As Long
Private Declare Function apiCreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Long) As Long
Private Declare Function apiDestroyWindow Lib "user32" Alias "DestroyWindow" (ByVal hWnd As Long) As Long
Private Declare Function apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function apiMoveWindow Lib "user32" Alias "MoveWindow" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function apiSendMessageLongLongLongLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Long) As Long
Private Declare Function apiSendMessageLongLongLongString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As String) As Long
Private Declare Function apiInitCommonControls Lib "Comctl32" Alias "InitCommonControls" () As Long
Private Declare Function apiInitCommonControlsEx Lib "Comctl32" Alias "InitCommonControlsEx" (ByRef iccex As tagInitCommonControlsEx) As Boolean
Private bInit   As Boolean
Private hAVI    As Long
Private hModule As Long
Friend Function PlayAviCtrl(ByVal hParent As Long, ByVal sFileName As String, Optional ResourceID As Long = 0, Optional bCenter As Boolean = True, Optional bAutoPlay As Boolean = True, Optional bTransparent As Boolean = True) As Boolean
   Dim lStyle As Long
   Dim ret    As Long
   StopAviCtrl
   On Error GoTo ErrPlay
   If Not (bInit = True) Then Call InitComctl32(ICC_ANIMATE_CLASS)
   lStyle = WS_CHILD Or WS_VISIBLE
   If bAutoPlay = True Then lStyle = lStyle Or ACS_AUTOPLAY
   If bTransparent = True Then lStyle = lStyle Or ACS_TRANSPARENT
   If hAVI <> 0 Then Call apiDestroyWindow(hAVI)
   If ResourceID = 0 Then
      hAVI = apiCreateWindowEx(0, ANIMATE_CLASSA, vbNullString, lStyle, 0, 0, 0, 0, hParent, 0&, App.hInstance, ByVal 0&)
      ret = apiSendMessageLongLongLongString(hAVI, ACM_OPEN, 0, ByVal sFileName)
   Else
      hModule = apiLoadLibrary(sFileName)
      If hModule = 0 Then GoTo ErrPlay
      hAVI = apiCreateWindowEx(0, ANIMATE_CLASSA, vbNullString, lStyle, 0, 0, 0, 0, hParent, 0&, hModule, ByVal 0&)
      ret = apiSendMessageLongLongLongLong(hAVI, ACM_OPEN, 0, ByVal ResourceID)
   End If
   If ret = 0 Or hAVI = 0 Then GoTo ErrPlay
   ' If bCenter = True Then Call CenterAVI(hParent)
   Call apiSendMessageLongLongLongLong(hAVI, ACM_PLAY, -1, ByVal 0&)
   PlayAviCtrl = True
   Exit Function
ErrPlay:
   If hModule <> 0 Then Call apiFreeLibrary(hModule)
   Call apiDestroyWindow(hAVI)
End Function
Friend Sub StopAviCtrl()
   On Error Resume Next
   If hAVI <> 0 Then
      Call apiSendMessageLongLongLongLong(hAVI, ACM_STOP, 0, ByVal 0&)
      Call apiSendMessageLongLongLongLong(hAVI, ACM_OPEN, 0, ByVal 0&)
      Call apiDestroyWindow(hAVI)
   End If
   If hModule <> 0 Then Call apiFreeLibrary(hModule)
End Sub
Private Sub InitComctl32(dwFlags As Long)
   On Error Resume Next
   Dim icc As tagInitCommonControlsEx
   bInit = True
   On Error GoTo Err_OldVersion
   icc.dwSize = Len(icc)
   icc.dwICC = dwFlags
   Call apiInitCommonControlsEx(icc)
   On Error GoTo 0
   Exit Sub
Err_OldVersion:
   Call apiInitCommonControls
End Sub
Private Sub CenterAVI(ByVal h As Long)
   On Error Resume Next
   Dim rcAVI    As RECT
   Dim rcParent As RECT
   Call apiGetWindowRect(h, rcParent)
   Call apiGetWindowRect(hAVI, rcAVI)
   Call apiMoveWindow(hAVI, (rcParent.right - rcParent.left - rcAVI.right + rcAVI.left) / 2, (rcParent.bottom - rcParent.top - rcAVI.bottom + rcAVI.top) / 2, rcAVI.right - rcAVI.left, rcAVI.bottom - rcAVI.top, True)
End Sub
