VERSION 5.00
Begin VB.Form frmSettings 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   9975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12975
   ControlBox      =   0   'False
   DrawWidth       =   3
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9975
   ScaleMode       =   0  'User
   ScaleWidth      =   12975
   ShowInTaskbar   =   0   'False
   Tag             =   "1"
   Begin VB.Frame fraMain 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   9915
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   12915
      Begin VB.Frame fracmdExit 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   840
         TabIndex        =   27
         Top             =   8400
         Width           =   1695
         Begin VB.Label lblExit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exit"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   600
            TabIndex        =   28
            Top             =   135
            Width           =   360
         End
      End
      Begin VB.Frame fracmdClose 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   10200
         TabIndex        =   25
         Top             =   8400
         Width           =   1695
         Begin VB.Label lblClose 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Close"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   480
            TabIndex        =   26
            Top             =   135
            Width           =   570
         End
      End
      Begin VB.Frame fracmdThumb 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1230
         Left            =   960
         TabIndex        =   20
         Top             =   3600
         Width           =   5000
         Begin VB.PictureBox picThumb 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   967
            Left            =   120
            Picture         =   "frmSettings.frx":000C
            ScaleHeight     =   960
            ScaleWidth      =   960
            TabIndex        =   23
            Top             =   120
            Width           =   967
         End
         Begin VB.PictureBox picThumbB 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   967
            Left            =   3240
            Picture         =   "frmSettings.frx":304E
            ScaleHeight     =   960
            ScaleWidth      =   960
            TabIndex        =   22
            Top             =   0
            Visible         =   0   'False
            Width           =   967
         End
         Begin VB.PictureBox picThumbG 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   967
            Left            =   4200
            Picture         =   "frmSettings.frx":6090
            ScaleHeight     =   960
            ScaleWidth      =   960
            TabIndex        =   21
            Top             =   0
            Visible         =   0   'False
            Width           =   967
         End
         Begin VB.Label lblThumb 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Thumb Stick calibration"
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   1680
            TabIndex        =   24
            Top             =   480
            Width           =   3255
         End
      End
      Begin VB.Frame fracmdGamepad 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1230
         Left            =   960
         TabIndex        =   15
         Top             =   5200
         Width           =   5000
         Begin VB.PictureBox picGamepad 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   967
            Left            =   120
            Picture         =   "frmSettings.frx":90D2
            ScaleHeight     =   960
            ScaleWidth      =   960
            TabIndex        =   18
            Top             =   120
            Width           =   967
         End
         Begin VB.PictureBox picGamepadG 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   967
            Left            =   4200
            Picture         =   "frmSettings.frx":C114
            ScaleHeight     =   960
            ScaleWidth      =   960
            TabIndex        =   17
            Top             =   0
            Visible         =   0   'False
            Width           =   967
         End
         Begin VB.PictureBox picGamepadB 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   967
            Left            =   3240
            Picture         =   "frmSettings.frx":F156
            ScaleHeight     =   960
            ScaleWidth      =   960
            TabIndex        =   16
            Top             =   0
            Visible         =   0   'False
            Width           =   967
         End
         Begin VB.Label lblGamepad 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Gamepad button mapping"
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   1680
            TabIndex        =   19
            Top             =   480
            Width           =   3255
         End
      End
      Begin VB.Frame fracmdPointer 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1230
         Left            =   960
         TabIndex        =   10
         Top             =   2000
         Width           =   5000
         Begin VB.PictureBox picPointer 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   967
            Left            =   120
            Picture         =   "frmSettings.frx":12198
            ScaleHeight     =   960
            ScaleWidth      =   960
            TabIndex        =   13
            Top             =   120
            Width           =   967
         End
         Begin VB.PictureBox picPointerG 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   967
            Left            =   4200
            Picture         =   "frmSettings.frx":151DA
            ScaleHeight     =   960
            ScaleWidth      =   960
            TabIndex        =   12
            Top             =   0
            Visible         =   0   'False
            Width           =   967
         End
         Begin VB.PictureBox picPointerB 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   967
            Left            =   3240
            Picture         =   "frmSettings.frx":1821C
            ScaleHeight     =   960
            ScaleWidth      =   960
            TabIndex        =   11
            Top             =   0
            Visible         =   0   'False
            Width           =   967
         End
         Begin VB.Label lblPointer 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Pointer/Point of view"
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   1680
            TabIndex        =   14
            Top             =   480
            Width           =   3255
         End
      End
      Begin VB.Frame fracmdDisplay 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1230
         Left            =   6720
         TabIndex        =   5
         Top             =   2000
         Width           =   5000
         Begin VB.PictureBox picDisplay 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   967
            Left            =   120
            Picture         =   "frmSettings.frx":1B25E
            ScaleHeight     =   960
            ScaleWidth      =   960
            TabIndex        =   8
            Top             =   120
            Width           =   967
         End
         Begin VB.PictureBox picDisplayG 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   967
            Left            =   4200
            Picture         =   "frmSettings.frx":1E2A0
            ScaleHeight     =   960
            ScaleWidth      =   960
            TabIndex        =   7
            Top             =   0
            Visible         =   0   'False
            Width           =   967
         End
         Begin VB.PictureBox picDisplayB 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   967
            Left            =   3240
            Picture         =   "frmSettings.frx":212E2
            ScaleHeight     =   960
            ScaleWidth      =   960
            TabIndex        =   6
            Top             =   0
            Visible         =   0   'False
            Width           =   967
         End
         Begin VB.Label lblDisplay 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Display"
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   1680
            TabIndex        =   9
            Top             =   480
            Width           =   3255
         End
      End
      Begin VB.Frame fracmdDirectx 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1230
         Left            =   6720
         TabIndex        =   3
         Top             =   5200
         Width           =   5000
         Begin VB.PictureBox picDirectxG 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   967
            Left            =   3600
            Picture         =   "frmSettings.frx":24324
            ScaleHeight     =   960
            ScaleWidth      =   960
            TabIndex        =   32
            Top             =   0
            Visible         =   0   'False
            Width           =   967
         End
         Begin VB.PictureBox picDirectxB 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   967
            Left            =   2520
            Picture         =   "frmSettings.frx":27366
            ScaleHeight     =   960
            ScaleWidth      =   960
            TabIndex        =   31
            Top             =   0
            Visible         =   0   'False
            Width           =   967
         End
         Begin VB.PictureBox picDirectx 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   967
            Left            =   120
            Picture         =   "frmSettings.frx":2A3A8
            ScaleHeight     =   960
            ScaleWidth      =   960
            TabIndex        =   30
            Top             =   120
            Width           =   967
         End
         Begin VB.Label lblDirectx 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Direct X"
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   1680
            TabIndex        =   4
            Top             =   480
            Width           =   3255
         End
      End
      Begin VB.Frame fracmdSound 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1230
         Left            =   6720
         TabIndex        =   1
         Top             =   3600
         Width           =   5000
         Begin VB.PictureBox picSoundG 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   967
            Left            =   3360
            Picture         =   "frmSettings.frx":2D3EA
            ScaleHeight     =   960
            ScaleWidth      =   960
            TabIndex        =   35
            Top             =   0
            Visible         =   0   'False
            Width           =   967
         End
         Begin VB.PictureBox picSoundB 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   967
            Left            =   2520
            Picture         =   "frmSettings.frx":3042C
            ScaleHeight     =   960
            ScaleWidth      =   960
            TabIndex        =   34
            Top             =   0
            Visible         =   0   'False
            Width           =   967
         End
         Begin VB.PictureBox picSound 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   967
            Left            =   120
            Picture         =   "frmSettings.frx":3346E
            ScaleHeight     =   960
            ScaleWidth      =   960
            TabIndex        =   33
            Top             =   120
            Width           =   967
         End
         Begin VB.Label lblSound 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Sound"
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   1680
            TabIndex        =   2
            Top             =   480
            Width           =   3255
         End
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " Settings"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   840
         TabIndex        =   29
         Top             =   480
         Width           =   11100
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private isloaded As Boolean
Private dragx As Long
Private dragy As Long
Private Sub Form_Load()
   On Error Resume Next
   CreateRoundRectFromWindow Me
   CreateRoundRectFromWindow fraMain
   CreateRoundRectFromWindow fracmdPointer
   CreateRoundRectFromWindow fracmdThumb
   CreateRoundRectFromWindow fracmdGamepad
   CreateRoundRectFromWindow fracmdSound
   CreateRoundRectFromWindow fracmdDisplay
   CreateRoundRectFromWindow fracmdDirectx
   dragx = -1
   dragy = -1
   Me.left = (Screen.Width - Me.Width) / 2
   Me.top = (Screen.Height - Me.Height) / 2
   isloaded = True
End Sub
Private Sub Form_Activate()
   WindowTransparency Me.hWnd, displayTransSettings, vbBlack
End Sub
Public Sub CenterForm(oForm As Object)
   On Error Resume Next
   oForm.Move (Screen.Width - oForm.Width) / 2, (Screen.Height - oForm.Height) / 2
End Sub
Private Sub fraMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If fracmdExit.BackColor <> &H404040 Then fracmdExit.BackColor = &H404040
   If fracmdClose.BackColor <> &H404040 Then fracmdClose.BackColor = &H404040
   If fracmdPointer.BackColor <> vbBlack Then fracmdPointer.BackColor = vbBlack: picPointer.Picture = picPointerB.Picture
   If fracmdThumb.BackColor <> vbBlack Then fracmdThumb.BackColor = vbBlack: picThumb.Picture = picThumbB.Picture
   If fracmdGamepad.BackColor <> vbBlack Then fracmdGamepad.BackColor = vbBlack: picGamepad.Picture = picGamepadB.Picture
   If fracmdSound.BackColor <> vbBlack Then fracmdSound.BackColor = vbBlack: picSound.Picture = picSoundB.Picture
   If fracmdDisplay.BackColor <> vbBlack Then fracmdDisplay.BackColor = vbBlack: picDisplay.Picture = picDisplayB.Picture
   If fracmdDirectx.BackColor <> vbBlack Then fracmdDirectx.BackColor = vbBlack: picDirectx.Picture = picDirectxB.Picture
   If dragx > -1 Then
      If x > dragx Then
         Me.left = Me.left + (x - dragx)
      ElseIf x < dragx Then
         Me.left = Me.left - (dragx - x)
      End If
   End If
   If dragy > -1 Then
      If y > dragy Then
         Me.top = Me.top + (y - dragy)
      ElseIf y < dragy Then
         Me.top = Me.top - (dragy - y)
      End If
   End If
End Sub
Private Sub fraMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   dragx = x
   dragy = y
End Sub
Private Sub fraMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   dragx = -1
   dragy = -1
End Sub
Private Sub lblClose_Click()
   fracmdClose_Click
End Sub
Private Sub lblExit_Click()
   fracmdExit_Click
End Sub
Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   fraMain_MouseDown Button, Shift, x, y
End Sub
Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   fraMain_MouseMove Button, Shift, x, y
End Sub
Private Sub lblTitle_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   fraMain_MouseUp Button, Shift, x, y
End Sub
Private Sub fracmdPointer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   fracmdPointer.BackColor = &H808080
   picPointer.Picture = picPointerG.Picture
End Sub
Private Sub fracmdThumb_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   fracmdThumb.BackColor = &H808080
   picThumb.Picture = picThumbG.Picture
End Sub
Private Sub fracmdGamepad_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   fracmdGamepad.BackColor = &H808080
   picGamepad.Picture = picGamepadG.Picture
End Sub
Private Sub fracmdSound_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   fracmdSound.BackColor = &H808080
   picSound.Picture = picSoundG.Picture
End Sub
Private Sub fracmdDisplay_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   fracmdDisplay.BackColor = &H808080
   picDisplay.Picture = picDisplayG.Picture
End Sub
Private Sub fracmdDirectx_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   fracmdDirectx.BackColor = &H808080
   picDirectx.Picture = picDirectxG.Picture
End Sub
Private Sub fracmdExit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   fracmdExit.BackColor = &H808080
End Sub
Private Sub fracmdClose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   fracmdClose.BackColor = &H808080
End Sub
Private Sub fracmdPointer_Click()
   On Error Resume Next
   If isloaded = False Then Exit Sub
   frmPointer.left = Me.left
   frmPointer.top = Me.top
   frmPointer.show
   frmMain.SetWindowPos frmPointer.hWnd, -1, 0, 0, 0, 0, False, False
   Unload Me
End Sub
Private Sub fracmdThumb_Click()
   On Error Resume Next
   If isloaded = False Then Exit Sub
   frmThumb.left = Me.left
   frmThumb.top = Me.top
   frmThumb.show
   frmMain.SetWindowPos frmThumb.hWnd, -1, 0, 0, 0, 0, False, False
   Unload Me
End Sub
Private Sub fracmdGamepad_Click()
   On Error Resume Next
   If isloaded = False Then Exit Sub
   frmGamepad.left = Me.left
   frmGamepad.top = Me.top
   frmGamepad.show
   frmMain.SetWindowPos frmGamepad.hWnd, -1, 0, 0, 0, 0, False, False
   Unload Me
End Sub
Private Sub fracmdSound_Click()
   On Error Resume Next
   If isloaded = False Then Exit Sub
   frmSound.left = Me.left
   frmSound.top = Me.top
   frmSound.show
   frmMain.SetWindowPos frmSound.hWnd, -1, 0, 0, 0, 0, False, False
   Unload Me
End Sub
Private Sub fracmdDisplay_Click()
   On Error Resume Next
   If isloaded = False Then Exit Sub
   frmDisplay.left = Me.left
   frmDisplay.top = Me.top
   frmDisplay.show
   frmMain.SetWindowPos frmDisplay.hWnd, -1, 0, 0, 0, 0, False, False
   Unload Me
End Sub
Private Sub fracmdDirectx_Click()
   If isloaded = False Then Exit Sub
   frmDirectX.left = Me.left
   frmDirectX.top = Me.top
   frmDirectX.show
   frmMain.SetWindowPos frmDirectX.hWnd, -1, 0, 0, 0, 0, False, False
   Unload Me
End Sub
Private Sub fracmdExit_Click()
   On Error Resume Next
   If isloaded = False Then Exit Sub
   frmCloseDialog.show
   frmMain.SetWindowPos frmCloseDialog.hWnd, -1, 0, 0, 0, 0, False, False
End Sub
Private Sub fracmdClose_Click()
   On Error Resume Next
   If isloaded = False Then Exit Sub
   Unload Me
End Sub
Private Sub lblDirectx_Click()
   fracmdDirectx_Click
End Sub
Private Sub lblDisplay_Click()
   fracmdDisplay_Click
End Sub
Private Sub lblGamepad_Click()
   fracmdGamepad_Click
End Sub
Private Sub lblPointer_Click()
   fracmdPointer_Click
End Sub
Private Sub lblSound_Click()
   fracmdSound_Click
End Sub
Private Sub lblThumb_Click()
   fracmdThumb_Click
End Sub
Private Sub picDirectx_Click()
   fracmdDirectx_Click
End Sub
Private Sub picDisplay_Click()
   fracmdDisplay_Click
End Sub
Private Sub picGamepad_Click()
   fracmdGamepad_Click
End Sub
Private Sub picPointer_Click()
   fracmdPointer_Click
End Sub
Private Sub picSound_Click()
   fracmdSound_Click
End Sub
Private Sub picThumb_Click()
   fracmdThumb_Click
End Sub
