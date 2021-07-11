Attribute VB_Name = "modWheel"
Option Explicit
Private Const GWL_WNDPROC        As Long = -4
Private Const WM_MOUSEWHEEL      As Long = &H20A
Private Const CB_GETDROPPEDSTATE As Long = &H157
Private Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByRef wParam As Any, ByRef lParam As Any) As Long
Private Declare Function apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function apiGetParent Lib "user32" Alias "GetParent" (ByVal hWnd As Long) As Long
Private Declare Function apiGetTickCount Lib "user32" Alias "GetTickCount" () As Long
Private lasttick  As Long
Public scrollposi As Long
Public Sub WheelHook(ByVal hWnd As Long)
   On Error Resume Next
   SetProp hWnd, "PrevWndProc", SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Public Sub WheelUnHook(ByVal hWnd As Long)
   On Error Resume Next
   SetWindowLong hWnd, GWL_WNDPROC, GetProp(hWnd, "PrevWndProc")
   RemoveProp hWnd, "PrevWndProc"
End Sub
Private Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   On Error Resume Next
   Dim Rotation As Long
   If Lmsg = WM_MOUSEWHEEL Then
      Rotation = wParam / 65536
      If Rotation = Abs(Rotation) Then
         scrollposi = Rotation
      Else
         scrollposi = Rotation
      End If
   Else
   End If
   WindowProc = CallWindowProc(GetProp(Lwnd, "PrevWndProc"), Lwnd, Lmsg, wParam, lParam)
End Function
Public Function IsOver(ByVal hWnd As Long, ByVal lX As Long, ByVal lY As Long) As Boolean
   Dim rectCtl As RECT
   apiGetWindowRect hWnd, rectCtl
   With rectCtl
      IsOver = (lX >= .left And lX <= .right And lY >= .top And lY <= .bottom)
   End With
End Function
Private Function GetForm(ByVal hWnd As Long) As Form
   For Each GetForm In Forms
      If GetForm.hWnd = hWnd Then Exit Function
   Next GetForm
   Set GetForm = Nothing
End Function
'Public Sub PictureBoxZoom(ByRef picBox As PictureBox, ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
'    picBox.Cls
'    picBox.Print "MouseWheel " & IIf(Rotation < 0, "Down", "Up")
'End Sub
'' Control Specific Behaviour
'' ================================================
'Public Sub FlexGridScroll(ByRef FG As MSFlexGrid, ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
'  Dim NewValue As Long
'  Dim Lstep As Single
'
'  On Error Resume Next
'  With FG
'    Lstep = .Height / .RowHeight(0)
'    Lstep = Int(Lstep)
'    If .Rows < Lstep Then Exit Sub
'    Do While Not (.RowIsVisible(.TopRow + Lstep))
'      Lstep = Lstep - 1
'    Loop
'    If Rotation > 0 Then
'        NewValue = .TopRow - Lstep
'        If NewValue < 1 Then
'            NewValue = 1
'        End If
'    Else
'        NewValue = .TopRow + Lstep
'        If NewValue > .Rows - 1 Then
'            NewValue = .Rows - 1
'        End If
'    End If
'    .TopRow = NewValue
'  End With
'End Sub
