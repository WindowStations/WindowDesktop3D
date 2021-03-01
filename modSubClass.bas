Attribute VB_Name = "modSubClass"
Option Explicit
Private Declare Function apiSetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function apiCallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function apiSetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function apiGetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_EXSTYLE          As Long = -20
Private Const GWL_WNDPROC          As Long = (-4)
Private Const WM_EXITSIZEMOVE      As Long = &H232
Private Const WM_MOUSEMOVE         As Long = &H200
Private Const WM_LBUTTONDOWN       As Long = &H201
Private Const WM_NCLBUTTONDOWN     As Long = &HA1
Private Const WM_SYSCOMMAND        As Long = &H112
Private Const WM_COMMAND           As Long = 273
Private Const WM_WINDOWPOSCHANGING As Long = 70
Private Const WM_WINDOWPOSCHANGED  As Long = 71
Private Const WM_MOVE              As Long = 3
Private Const WM_SIZE              As Long = 5
Private Const WM_ACTIVATE          As Long = 6
Private Const WM_SIZING            As Long = 532
Private Const WM_MOVING            As Long = 534
Private Const WM_MOUSEACTIVATE     As Long = 33
Private Const SWP_NOSIZE           As Long = &H1
Private Const SWP_NOMOVE           As Long = &H2
Private Const SWP_NOACTIVATE       As Long = &H10
Private Const SWP_SHOWWINDOW       As Long = &H40
Private Const HWND_BOTTOM          As Long = 1
Private Const WS_EX_NOACTIVATE     As Long = &H8000000
Public glPrevWndProc               As Long

Public apb As New clsAVIToPictureBox
Public Function fSubClass(ByVal hWnd As Long) As Long ' NOTE: Comment all calls to fSubClass while in IDE.
    fSubClass = apiSetWindowLong(hWnd, GWL_WNDPROC, AddressOf pMyWindowProc)
End Function
Public Sub pUnSubClass(ByVal hWnd As Long)
    Call apiSetWindowLong(hWnd, GWL_WNDPROC, glPrevWndProc)
End Sub
Public Function pMyWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    If uMsg = WM_MOUSEACTIVATE Or uMsg = WM_ACTIVATE Then
'        Call apiSetWindowLong(frmWallpaper.hWnd, GWL_EXSTYLE, apiGetWindowLong(frmWallpaper.hWnd, GWL_EXSTYLE) Or WS_EX_NOACTIVATE)
'        Exit Function
'    End If
'    If uMsg = WM_SYSCOMMAND Then
'        Exit Function
'    End If
    If uMsg = WM_WINDOWPOSCHANGED Then
        'Call apiSetWindowPos(frmWallpaper.hWnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE)
        frmMain.wpc.WindowPosChanged
        
       ' Exit Function
    End If
    If uMsg = WM_WINDOWPOSCHANGING Then
       ' Exit Function
       ' Call apiSetWindowPos(frmWallpaper.hWnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE)
        
    End If
    pMyWindowProc = apiCallWindowProc(glPrevWndProc, hw, uMsg, wParam, lParam)
End Function
