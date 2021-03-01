VERSION 5.00
Begin VB.Form frmGamepad 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Gamepad button mapping"
   ClientHeight    =   9975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12975
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   12975
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   7200
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   12
      Top             =   8520
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   9735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12735
      Begin VB.Frame frachkDisableGamepad 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   840
         TabIndex        =   19
         Tag             =   "1,20"
         Top             =   1320
         Width           =   5000
         Begin VB.CheckBox chkDisableGamepad 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   0
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   360
            Width           =   200
         End
         Begin VB.Label lblDisablegamepad 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Disable gamepad event handling"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   360
            TabIndex        =   21
            Top             =   360
            Width           =   3405
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5415
         Left            =   6960
         TabIndex        =   16
         Tag             =   "1,20"
         Top             =   2400
         Width           =   5535
         Begin VB.ListBox List2 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   4755
            ItemData        =   "frmGamepad.frx":0000
            Left            =   0
            List            =   "frmGamepad.frx":0031
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   480
            Width           =   5000
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Simulated event"
            ForeColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   1680
         End
      End
      Begin VB.Frame fratra1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5415
         Left            =   840
         TabIndex        =   13
         Tag             =   "1,20"
         Top             =   2400
         Width           =   5535
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   4755
            ItemData        =   "frmGamepad.frx":0191
            Left            =   0
            List            =   "frmGamepad.frx":01BF
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   480
            Width           =   5000
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gamepad input "
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   1665
         End
      End
      Begin VB.Frame Frame5 
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
         Left            =   4680
         TabIndex        =   10
         Top             =   8400
         Width           =   1695
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default all"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   240
            TabIndex        =   11
            Top             =   135
            Width           =   1050
         End
      End
      Begin VB.Frame Frame4 
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
         Left            =   2760
         TabIndex        =   8
         Top             =   8400
         Width           =   1695
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default input"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   120
            TabIndex        =   9
            Top             =   135
            Width           =   1350
         End
      End
      Begin VB.Frame Frame3 
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
         Left            =   8280
         TabIndex        =   6
         Top             =   8400
         Width           =   1695
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apply"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   480
            TabIndex        =   7
            Top             =   135
            Width           =   600
         End
      End
      Begin VB.Frame Frame8 
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
         TabIndex        =   2
         Top             =   8400
         Width           =   1695
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Back"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   600
            TabIndex        =   3
            Top             =   135
            Width           =   480
         End
      End
      Begin VB.Frame Frame2 
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
         TabIndex        =   4
         Top             =   8400
         Width           =   1695
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Close"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   480
            TabIndex        =   5
            Top             =   135
            Width           =   570
         End
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Gamepad button mapping"
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
         TabIndex        =   1
         Top             =   480
         Width           =   11100
      End
   End
End
Attribute VB_Name = "frmGamepad"
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
    CreateRoundRectFromWindow Frame1
    chkDisableGamepad.Value = keymapDisablegamepad
    dragx = -1
    dragy = -1
    isloaded = True
End Sub
Private Sub Form_Activate()
    WindowTransparency Me.hWnd, displayTransSettings, vbBlack
End Sub
Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    dragx = x
    dragy = y
End Sub
Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    dragx = -1
    dragy = -1
End Sub
Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Frame1_MouseDown Button, Shift, x, y
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Frame1_MouseMove Button, Shift, x, y
End Sub
Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Frame1_MouseUp Button, Shift, x, y
End Sub

Private Sub Frame8_Click()
    If isloaded = False Then Exit Sub
    frmSettings.show
    frmSettings.top = Me.top
    frmSettings.left = Me.left
    frmMain.SetWindowPos frmSettings.hWnd, -1, 0, 0, 0, 0, False, False
    Unload Me
End Sub
Private Sub Label8_Click()
    If isloaded = False Then Exit Sub
    Frame8_Click
End Sub
Private Sub Frame4_Click()
    On Error Resume Next
    If isloaded = False Then Exit Sub
    If List1.ListIndex = 0 Then keymapA = "1": SaveSetting "Window3D", "ButtonMap", "AButton", 1
    If List1.ListIndex = 1 Then keymapMenu = "2": SaveSetting "Window3D", "ButtonMap", "Menu", 2
    If List1.ListIndex = 2 Then keymapB = "3": SaveSetting "Window3D", "ButtonMap", "BButton", 3
    If List1.ListIndex = 3 Then keymapY = "4": SaveSetting "Window3D", "ButtonMap", "YButton", 4
    If List1.ListIndex = 4 Then keymapX = "5": SaveSetting "Window3D", "ButtonMap", "XButton", 5
    If List1.ListIndex = 5 Then keymapLeftBumper = "6": SaveSetting "Window3D", "ButtonMap", "LeftBumper", 6
    If List1.ListIndex = 6 Then keymapRightBumper = "7": SaveSetting "Window3D", "ButtonMap", "RightBumper", 7
    If List1.ListIndex = 7 Then keymapLeftStick = "8": SaveSetting "Window3D", "ButtonMap", "LeftStick", 8
    If List1.ListIndex = 8 Then keymapRightStick = "9": SaveSetting "Window3D", "ButtonMap", "RightStick", 9
    If List1.ListIndex = 9 Then keymapDLeft = "10": SaveSetting "Window3D", "ButtonMap", "DLeft", 10
    If List1.ListIndex = 10 Then keymapDRight = "11": SaveSetting "Window3D", "ButtonMap", "DRight", 11
    If List1.ListIndex = 11 Then keymapDUp = "12": SaveSetting "Window3D", "ButtonMap", "DUp", 12
    If List1.ListIndex = 12 Then keymapDDown = "13": SaveSetting "Window3D", "ButtonMap", "DDown", 13
    If List1.ListIndex = 13 Then keymapChange = "14": SaveSetting "Window3D", "ButtonMap", "Change", 14
    If List1.ListIndex > -1 Then Frame8_Click
End Sub
Private Sub Label6_Click()
    If isloaded = False Then Exit Sub
    Frame4_Click
End Sub
Private Sub Frame5_Click()
    On Error Resume Next
    If isloaded = False Then Exit Sub
    keymapA = "1": SaveSetting "Window3D", "ButtonMap", "AButton", 1
    keymapMenu = "2": SaveSetting "Window3D", "ButtonMap", "Menu", 2
    keymapB = "3": SaveSetting "Window3D", "ButtonMap", "BButton", 3
    keymapY = "4": SaveSetting "Window3D", "ButtonMap", "YButton", 4
    keymapX = "5": SaveSetting "Window3D", "ButtonMap", "XButton", 5
    keymapLeftBumper = "6": SaveSetting "Window3D", "ButtonMap", "LeftBumper", 6
    keymapRightBumper = "7": SaveSetting "Window3D", "ButtonMap", "RightBumper", 7
    keymapLeftStick = "8": SaveSetting "Window3D", "ButtonMap", "LeftStick", 8
    keymapRightStick = "9": SaveSetting "Window3D", "ButtonMap", "RightStick", 9
    keymapDLeft = "10": SaveSetting "Window3D", "ButtonMap", "DLeft", 10
    keymapDRight = "11": SaveSetting "Window3D", "ButtonMap", "DRight", 11
    keymapDUp = "12": SaveSetting "Window3D", "ButtonMap", "DUp", 12
    keymapDDown = "13": SaveSetting "Window3D", "ButtonMap", "DDown", 13
    keymapChange = "14": SaveSetting "Window3D", "ButtonMap", "Change", 14
    Frame8_Click
End Sub
Private Sub Label7_Click()
    If isloaded = False Then Exit Sub
    Frame5_Click
End Sub

Private Sub chkDisableGamepad_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If isloaded = False Then Exit Sub
    CheckBoxSetting
End Sub
Private Sub lblDisablegamepad_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If chkDisableGamepad.Value = vbUnchecked Then
        chkDisableGamepad.Value = vbChecked
    Else
        chkDisableGamepad.Value = vbUnchecked
    End If
    CheckBoxSetting
End Sub
Private Sub CheckBoxSetting()
    keymapDisablegamepad = chkDisableGamepad.Value
    If keymapDisablegamepad = 1 Then
        frmMain.xinputClass.Disable
    Else
        frmMain.xinputClass.Enable
    End If
End Sub

Private Sub Frame3_Click()
    On Error Resume Next
    If isloaded = False Then Exit Sub
    
    keymapDisablegamepad = chkDisableGamepad.Value
    SaveSetting "Window3D", "ButtonMap", "DisableGamepad", CStr(keymapDisablegamepad)
    
    If List1.ListIndex <> -1 And List2.ListIndex <> -1 Then
        If List1.List(List1.ListIndex) = "A button" Then SaveSetting "Window3D", "ButtonMap", "AButton", List2.ListIndex + 1: keymapA = CStr(List2.ListIndex + 1)
        If List1.List(List1.ListIndex) = "Menu" Then SaveSetting "Window3D", "ButtonMap", "Menu", List2.ListIndex + 1: keymapMenu = CStr(List2.ListIndex + 1)
        If List1.List(List1.ListIndex) = "B button" Then SaveSetting "Window3D", "ButtonMap", "BButton", List2.ListIndex + 1: keymapB = CStr(List2.ListIndex + 1)
        If List1.List(List1.ListIndex) = "Y button" Then SaveSetting "Window3D", "ButtonMap", "YButton", List2.ListIndex + 1: keymapY = CStr(List2.ListIndex + 1)
        If List1.List(List1.ListIndex) = "X button" Then SaveSetting "Window3D", "ButtonMap", "XButton", List2.ListIndex + 1: keymapX = CStr(List2.ListIndex + 1)
        If List1.List(List1.ListIndex) = "Left bumper" Then SaveSetting "Window3D", "ButtonMap", "LeftBumper", List2.ListIndex + 1: keymapLeftBumper = CStr(List2.ListIndex + 1)
        If List1.List(List1.ListIndex) = "Right bumper" Then SaveSetting "Window3D", "ButtonMap", "RightBumper", List2.ListIndex + 1: keymapRightBumper = CStr(List2.ListIndex + 1)
        If List1.List(List1.ListIndex) = "Left stick click" Then SaveSetting "Window3D", "ButtonMap", "LeftStick", List2.ListIndex + 1: keymapLeftStick = CStr(List2.ListIndex + 1)
        If List1.List(List1.ListIndex) = "Right stick click" Then SaveSetting "Window3D", "ButtonMap", "RightStick", List2.ListIndex + 1: keymapRightStick = CStr(List2.ListIndex + 1)
        If List1.List(List1.ListIndex) = "D-pad left" Then SaveSetting "Window3D", "ButtonMap", "DLeft", List2.ListIndex + 1: keymapDLeft = CStr(List2.ListIndex + 1)
        If List1.List(List1.ListIndex) = "D-pad right" Then SaveSetting "Window3D", "ButtonMap", "DRight", List2.ListIndex + 1: keymapDRight = CStr(List2.ListIndex + 1)
        If List1.List(List1.ListIndex) = "D-pad up" Then SaveSetting "Window3D", "ButtonMap", "DUp", List2.ListIndex + 1: keymapDUp = CStr(List2.ListIndex + 1)
        If List1.List(List1.ListIndex) = "D-pad down" Then SaveSetting "Window3D", "ButtonMap", "DDown", List2.ListIndex + 1: keymapDDown = CStr(List2.ListIndex + 1)
        If List1.List(List1.ListIndex) = "Change Window" Then SaveSetting "Window3D", "ButtonMap", "Change", List2.ListIndex + 1: keymapChange = CStr(List2.ListIndex + 1)
    End If
    Beep
End Sub
Private Sub Label2_Click()
    If isloaded = False Then Exit Sub
    Frame3_Click
End Sub
Private Sub Frame2_Click()
    If isloaded = False Then Exit Sub
    Unload Me
End Sub
Private Sub Label1_Click()
    If isloaded = False Then Exit Sub
    Frame2_Click
End Sub
Private Sub Frame8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Frame8.BackColor = &H808080
End Sub
Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Frame4.BackColor = &H808080
End Sub
Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Frame5.BackColor = &H808080
End Sub
Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Frame3.BackColor = &H808080
End Sub
Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Frame2.BackColor = &H808080
End Sub
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Frame8.BackColor <> &H404040 Then Frame8.BackColor = &H404040
    If Frame5.BackColor <> &H404040 Then Frame5.BackColor = &H404040
    If Frame4.BackColor <> &H404040 Then Frame4.BackColor = &H404040
    If Frame3.BackColor <> &H404040 Then Frame3.BackColor = &H404040
    If Frame2.BackColor <> &H404040 Then Frame2.BackColor = &H404040
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
'Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
'   Dim inx As Long
'   inx = List1.ListIndex + 1
'   If inx < 1 Then Exit Sub
'   Picture1.Picture = ImageList1.ListImages.Item(inx).Picture
'End Sub
'
'Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Dim inx As Long
'   inx = List1.ListIndex + 1
'   If inx < 1 Then Exit Sub
'   Picture1.Picture = ImageList1.ListImages.Item(inx).Picture
'End Sub
Private Sub lblDisablegamepad_Click()

End Sub


