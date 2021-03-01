VERSION 5.00
Begin VB.Form frmForm 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   Caption         =   "Text Display"
   ClientHeight    =   9975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   12975
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraMain 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12735
      Begin VB.Frame fraTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1300
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   12735
         Begin VB.Label lblTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Title"
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
            Height          =   705
            Left            =   0
            TabIndex        =   30
            Top             =   300
            Width           =   12735
         End
      End
      Begin VB.Frame frascrMain 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   6500
         Left            =   12360
         TabIndex        =   28
         Top             =   1320
         Width           =   240
      End
      Begin VB.Frame fracmdClose 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   10200
         TabIndex        =   1
         Top             =   8400
         Width           =   1695
         Begin VB.Label lblClose 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Close"
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
            Height          =   315
            Left            =   480
            TabIndex        =   2
            Top             =   135
            Width           =   570
         End
      End
      Begin VB.Frame fracmdBack 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   840
         TabIndex        =   3
         Top             =   8400
         Width           =   1695
         Begin VB.Label lblBack 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Back"
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
            Height          =   315
            Left            =   600
            TabIndex        =   4
            Top             =   135
            Width           =   480
         End
      End
      Begin VB.Frame fracmdApply 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   8280
         TabIndex        =   5
         Top             =   8400
         Width           =   1695
         Begin VB.Label lblApply 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apply"
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
            Height          =   315
            Left            =   480
            TabIndex        =   6
            Top             =   135
            Width           =   600
         End
      End
      Begin VB.Frame fraMainScroll 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6500
         Left            =   840
         TabIndex        =   7
         Top             =   1300
         Width           =   11055
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1215
            Left            =   120
            TabIndex        =   25
            Top             =   1680
            Width           =   5000
            Begin VB.PictureBox Picture1 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               ForeColor       =   &H80000008&
               Height          =   840
               Left            =   120
               ScaleHeight     =   810
               ScaleWidth      =   930
               TabIndex        =   27
               Top             =   240
               Width           =   967
            End
            Begin VB.CheckBox Check1 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               Caption         =   "Checkbox"
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
               Height          =   855
               Left            =   1320
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   240
               Width           =   3495
            End
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1215
            Left            =   0
            TabIndex        =   22
            Top             =   240
            Width           =   5000
            Begin VB.OptionButton Option1 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               Caption         =   "Option1"
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
               Height          =   495
               Left            =   240
               MaskColor       =   &H00000000&
               TabIndex        =   24
               Top             =   360
               Width           =   1575
            End
            Begin VB.OptionButton Option2 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               Caption         =   "Option2"
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
               Height          =   495
               Left            =   1920
               MaskColor       =   &H00000000&
               TabIndex        =   23
               Top             =   360
               Width           =   1575
            End
         End
         Begin VB.Frame fracmd1 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1230
            Left            =   120
            TabIndex        =   17
            Top             =   4680
            Width           =   5000
            Begin VB.PictureBox Picture10 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   967
               Left            =   3000
               Picture         =   "frmForm.frx":0000
               ScaleHeight     =   960
               ScaleWidth      =   960
               TabIndex        =   20
               Top             =   0
               Visible         =   0   'False
               Width           =   967
            End
            Begin VB.PictureBox Picture9 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   967
               Left            =   3960
               Picture         =   "frmForm.frx":3042
               ScaleHeight     =   960
               ScaleWidth      =   960
               TabIndex        =   19
               Top             =   0
               Visible         =   0   'False
               Width           =   967
            End
            Begin VB.PictureBox Picture2 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   967
               Left            =   120
               Picture         =   "frmForm.frx":6084
               ScaleHeight     =   960
               ScaleWidth      =   960
               TabIndex        =   18
               Top             =   120
               Width           =   967
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Gamepad button mapping"
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
               Height          =   495
               Left            =   1320
               TabIndex        =   21
               Top             =   480
               Width           =   3495
            End
         End
         Begin VB.Frame fralst1 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3135
            Left            =   5880
            TabIndex        =   13
            Tag             =   "1,20"
            Top             =   240
            Width           =   5000
            Begin VB.ListBox lst1 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
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
               Height          =   2550
               ItemData        =   "frmForm.frx":90C6
               Left            =   0
               List            =   "frmForm.frx":90D3
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   480
               Width           =   5000
            End
            Begin VB.Label lbllstTitle1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "List title"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00E0E0E0&
               Height          =   315
               Left            =   0
               TabIndex        =   16
               Top             =   0
               Width           =   810
            End
            Begin VB.Label Label1 
               BackColor       =   &H00C0C0C0&
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   15
               TabIndex        =   15
               Top             =   560
               Width           =   15
            End
         End
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6240
            TabIndex        =   12
            TabStop         =   0   'False
            Text            =   "1"
            Top             =   4200
            Width           =   1080
         End
         Begin VB.Frame fratra1 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   975
            Left            =   120
            TabIndex        =   8
            Tag             =   "1,20"
            Top             =   3360
            Width           =   5000
            Begin VB.Label lblSliderTitle1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Trackbar title"
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
               Height          =   315
               Left            =   0
               TabIndex        =   11
               Top             =   0
               Width           =   1350
            End
            Begin VB.Label lbltraValue1 
               BackColor       =   &H00C0C0C0&
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
               Height          =   240
               Left            =   15
               TabIndex        =   10
               Top             =   560
               Width           =   15
            End
            Begin VB.Label lbltra1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||"
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
               Height          =   315
               Left            =   0
               TabIndex        =   9
               Top             =   480
               Width           =   5160
            End
         End
      End
   End
End
Attribute VB_Name = "frmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private isloaded  As Boolean
Private dragx As Long
Private dragy As Long
Private dragscr As Boolean
Private Declare Function apiBringWindowToTop Lib "user32" Alias "BringWindowToTop" (ByVal hWnd As Long) As Long
Private Sub Form_Load()
    On Error Resume Next
    CreateRoundRectFromWindow Me
    CreateRoundRectFromWindow fraMain
    dragx = -1
    dragy = -1
    isloaded = True
End Sub
Private Sub Form_Activate()
    WindowTransparency Me.hWnd, 235, vbBlack
End Sub
Private Sub fraMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    dragx = x
    dragy = y
End Sub
Private Sub fraMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    dragx = -1
    dragy = -1
End Sub
Private Sub fraTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraMain_MouseDown Button, Shift, x, y
End Sub
Private Sub fraTitle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraMain_MouseMove Button, Shift, x, y
End Sub
Private Sub fraTitle_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraMain_MouseUp Button, Shift, x, y
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
Private Sub fraMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If fracmdBack.BackColor <> &H404040 Then fracmdBack.BackColor = &H404040
    If fracmdApply.BackColor <> &H404040 Then fracmdApply.BackColor = &H404040
    If fracmdClose.BackColor <> &H404040 Then fracmdClose.BackColor = &H404040
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
Private Sub fracmdBack_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fracmdBack.BackColor = &H808080
End Sub
Private Sub fracmdApply_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fracmdApply.BackColor = &H808080
End Sub
Private Sub fracmdClose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fracmdClose.BackColor = &H808080
End Sub
Private Sub vscrChange(ByVal x As Single, ByVal y As Single)
    Dim he As Long
    Dim ra As Double
    Dim tp As Long
    he = frascrMain.Height - 135
    ra = (y - 135) / he
    tp = 1000 - (ra * (fraMainScroll.Height - 4000))
    If Abs(tp) > (fraMainScroll.Height - 4000) Then
        tp = -(Abs(tp) - (fraMainScroll.Height - 4000))
        Exit Sub
    End If
    fraMainScroll.top = tp
End Sub
Private Sub frascrMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    dragscr = True
    If y < 135 Then Exit Sub
    If y > frascrMain.Height Then Exit Sub
    vscrChange x, y
End Sub
Private Sub frascrMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If dragscr = False Then Exit Sub
    If y < 135 Then Exit Sub
    If y > frascrMain.Height Then Exit Sub
    vscrChange x, y
End Sub
Private Sub frascrMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    dragscr = False
End Sub
Private Sub fracmdBack_Click()
    frmOptions.show
    frmOptions.top = Me.top
    frmOptions.left = Me.left
    frmMain.SetWindowPos frmOptions.hWnd, -1, 0, 0, 0, 0, False, False
    Unload Me
End Sub
Private Sub lblBack_Click()
    fracmdBack_Click
End Sub
Private Sub fracmdApply_Click()
    '    pointerTextSize = traTextSize.Value
    '    pointerTextPosition = traTextPosition.Value
    '    pointerTextSpeed = traTextSpeed.Value
    '    pointerTextFade = traTextFade.Value
    '    SaveSetting "Window3D", "ButtonMap", "TextSize", CStr(pointerTextSize)
    '    SaveSetting "Window3D", "ButtonMap", "TextPosition", CStr(pointerTextPosition)
    '    SaveSetting "Window3D", "ButtonMap", "TextSpeed", CStr(pointerTextSpeed)
    '    SaveSetting "Window3D", "ButtonMap", "TextFade", CStr(pointerTextFade)
    Beep
End Sub
Private Sub lblApply_Click()
    fracmdApply_Click
End Sub
Private Sub fracmdClose_Click()
    Unload Me
End Sub
Private Sub lblClose_Click()
    fracmdClose_Click
End Sub
