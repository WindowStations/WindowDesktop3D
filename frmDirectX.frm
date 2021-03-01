VERSION 5.00
Begin VB.Form frmDirectX 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
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
   Begin VB.Frame Frame1 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   9735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12735
      Begin VB.Frame fracmdspawn 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4680
         TabIndex        =   33
         Top             =   8400
         Width           =   1695
         Begin VB.Label lblRespawn 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Respawn"
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
            Left            =   360
            TabIndex        =   34
            Top             =   135
            Width           =   945
         End
      End
      Begin VB.Frame frachkVsync 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   6360
         TabIndex        =   30
         Tag             =   "1,20"
         Top             =   3000
         Width           =   5000
         Begin VB.CheckBox chkVSync 
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
            Height          =   315
            Left            =   0
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   240
            Width           =   200
         End
         Begin VB.Label lblvsync 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "V-Sync"
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
            TabIndex        =   31
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   10200
         TabIndex        =   27
         Top             =   8400
         Width           =   1695
         Begin VB.Label Label2 
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
            TabIndex        =   28
            Top             =   135
            Width           =   570
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   840
         TabIndex        =   25
         Top             =   8400
         Width           =   1695
         Begin VB.Label Label3 
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
            TabIndex        =   26
            Top             =   135
            Width           =   480
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   8280
         TabIndex        =   23
         Top             =   8400
         Width           =   1695
         Begin VB.Label Label4 
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
            TabIndex        =   24
            Top             =   135
            Width           =   600
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   6120
         TabIndex        =   18
         Top             =   1920
         Width           =   5655
         Begin VB.OptionButton Option1 
            BackColor       =   &H00000000&
            Caption         =   "Bi-linear"
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
            Left            =   600
            TabIndex        =   21
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00000000&
            Caption         =   "Tri-linear"
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
            Left            =   2160
            TabIndex        =   20
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00000000&
            Caption         =   "Anisotropic"
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
            Left            =   3600
            TabIndex        =   19
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Texture filter"
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
            Left            =   240
            TabIndex        =   22
            Top             =   0
            Width           =   1305
         End
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   960
         TabIndex        =   15
         Top             =   3720
         Width           =   5000
         Begin VB.TextBox txtQuant 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
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
            Left            =   3240
            TabIndex        =   16
            Text            =   "200"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quant cycle denominator 1/d"
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
            TabIndex        =   17
            Top             =   240
            Width           =   3045
         End
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   960
         TabIndex        =   12
         Top             =   2760
         Width           =   5000
         Begin VB.TextBox txtGravity 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
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
            Left            =   3240
            TabIndex        =   13
            Text            =   "-1000"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gravity"
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
            TabIndex        =   14
            Top             =   240
            Width           =   750
         End
      End
      Begin VB.Frame Frame11 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "TexFIndex"
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   960
         TabIndex        =   9
         Top             =   1800
         Width           =   5000
         Begin VB.TextBox txtTexFIndex 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
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
            Left            =   3240
            TabIndex        =   10
            Text            =   "4"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Anisotropy TexFIndex"
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
            Top             =   240
            Width           =   2220
         End
      End
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "TexFIndex"
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   960
         TabIndex        =   6
         Top             =   4680
         Width           =   5000
         Begin VB.TextBox txtFovY 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
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
            Left            =   3240
            TabIndex        =   7
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Field of view FovY"
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
            TabIndex        =   8
            Top             =   240
            Width           =   1875
         End
      End
      Begin VB.Frame Frame12 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "TexFIndex"
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   960
         TabIndex        =   3
         Top             =   5640
         Width           =   5000
         Begin VB.TextBox txtAspect 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
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
            Left            =   3240
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aspect ratio"
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
            TabIndex        =   5
            Top             =   240
            Width           =   1230
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2760
         TabIndex        =   1
         Top             =   8400
         Width           =   1695
         Begin VB.Label lblDefault 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default"
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
            Width           =   750
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "                               DirectX"
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
         Left            =   720
         TabIndex        =   29
         Top             =   480
         Width           =   11100
      End
   End
End
Attribute VB_Name = "frmDirectX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private isloaded  As Boolean
Private dragx As Long
Private dragy As Long


Private Sub Form_Load()
    On Error Resume Next
    CreateRoundRectFromWindow Me
    CreateRoundRectFromWindow Frame1
    
    
    If directxTexFilters = TextureFilter.TextureFilter_BiLinear Then
        Option1.Value = True
        Option2.Value = False
        Option3.Value = False
    ElseIf directxTexFilters = TextureFilter.TextureFilter_TriLinear Then
        Option2.Value = True
        Option1.Value = False
        Option3.Value = False
    ElseIf directxTexFilters = TextureFilter.TextureFilter_Anisotropic Then
        Option3.Value = True
        Option1.Value = False
        Option2.Value = False
    End If
    chkVSync.Value = Abs(CLng(directxVSync))
    txtQuant.Text = CStr(1 / directxQuant)
    txtGravity.Text = CStr(1 / directxGravity)
    txtTexFIndex.Text = CStr(directxTexFIndex)
    txtFovY.Text = CStr(directxFovY)
    txtAspect.Text = CStr(directxAspect)
    dragx = -1
    dragy = -1
    isloaded = True
End Sub
Private Sub Form_Activate()
    WindowTransparency Me.hWnd, displayTransSettings, vbBlack
End Sub






Private Sub fracmdSpawn_Click()
    frmMain.Respawn
End Sub




Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    dragx = x
    dragy = y
End Sub
Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    dragx = -1
    dragy = -1
End Sub


Private Sub Frame6_Click()
    directxTexFilters = 2
    directxVSync = False
    directxQuant = 1 / 200
    directxGravity = 1 / -1000
    directxTexFIndex = 4
    directxAnisotropy = 2 ^ (directxTexFIndex - 1)
    Frame3_Click
End Sub
Private Sub lblDefault_Click()
    Frame6_Click
End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Frame1_MouseDown Button, Shift, x, y
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Frame1_MouseMove Button, Shift, x, y
End Sub
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Frame1_MouseUp Button, Shift, x, y
End Sub
Private Sub Label2_Click()
Frame4_Click
End Sub

Private Sub Label3_Click()
Frame2_Click
End Sub

Private Sub Label4_Click()
Frame3_Click
End Sub
Private Sub Frame4_Click()
    Unload Me
End Sub
Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Frame4.BackColor = &H808080
End Sub
Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Frame3.BackColor = &H808080
End Sub
Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Frame2.BackColor = &H808080
End Sub
Private Sub fracmdspawn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fracmdspawn.BackColor = &H808080
End Sub
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Frame4.BackColor <> &H404040 Then Frame4.BackColor = &H404040
    If Frame3.BackColor <> &H404040 Then Frame3.BackColor = &H404040
    If Frame2.BackColor <> &H404040 Then Frame2.BackColor = &H404040
      If fracmdspawn.BackColor <> &H404040 Then fracmdspawn.BackColor = &H404040
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
Private Sub chkVSync_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If isloaded = False Then Exit Sub
   CheckBoxSetting
End Sub

Private Sub lblRespawn_Click()
fracmdSpawn_Click
End Sub

Private Sub lblvsync_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If isloaded = False Then Exit Sub
  If chkVSync.Value = vbUnchecked Then
     chkVSync.Value = vbChecked
  Else
     chkVSync.Value = vbUnchecked
  End If
  CheckBoxSetting
End Sub
Private Sub CheckBoxSetting()
    directxVSync = CBool(chkVSync.Value)
End Sub


Private Sub Option1_Click()
    If isloaded = False Then Exit Sub
    directxTexFilters = TextureFilter.TextureFilter_BiLinear
End Sub
Private Sub Option2_Click()
    If isloaded = False Then Exit Sub
    directxTexFilters = TextureFilter.TextureFilter_TriLinear
End Sub
Private Sub Option3_Click()
    If isloaded = False Then Exit Sub
    directxTexFilters = TextureFilter.TextureFilter_Anisotropic
End Sub
Private Sub txtQuant_Change()
    If isloaded = False Then Exit Sub
    If IsNumeric(txtQuant.Text) = True Then directxQuant = 1 / CDbl(txtQuant.Text)
End Sub
Private Sub txtGravity_Change()
    If isloaded = False Then Exit Sub
    If IsNumeric(txtGravity.Text) = True Then directxGravity = 1 / CSng(txtGravity.Text)
End Sub
'Private Sub txtAnisotropy_Change()
'    If isloaded = False Then Exit Sub
'    If IsNumeric(txtAnisotropy.Text) = True Then
'       directxAnisotropy = CLng(txtAnisotropy.Text)
'      'directxTexFIndex=   2 ^ (directxTexFIndex - 1)
'    End If
'End Sub
Private Sub txtTexFIndex_Change()
    If isloaded = False Then Exit Sub
    If IsNumeric(txtTexFIndex.Text) = True Then
        directxTexFIndex = CLng(txtTexFIndex.Text)
        directxAnisotropy = 2 ^ (directxTexFIndex - 1)
    End If
End Sub
Private Sub txtFovY_Change()
    If isloaded = False Then Exit Sub
    If IsNumeric(txtFovY.Text) = True Then directxFovY = CSng(txtFovY.Text)
End Sub
Private Sub txtAspect_Change()
    If isloaded = False Then Exit Sub
    If IsNumeric(txtAspect.Text) = True Then directxAspect = CSng(txtAspect.Text)
End Sub
Private Sub Frame2_Click()
    frmSettings.show
    frmSettings.top = Me.top
    frmSettings.left = Me.left
    frmMain.SetWindowPos frmSettings.hWnd, -1, 0, 0, 0, 0, False, False
    Unload Me
End Sub
Private Sub Frame3_Click()
    On Error Resume Next
    SaveSetting "Window3D", "DirectX", "TexFilters", CStr(directxTexFilters)
    SaveSetting "Window3D", "DirectX", "VSync", CStr(Abs(CLng(directxVSync)))
    SaveSetting "Window3D", "DirectX", "Quant", CStr(1 / directxQuant)
    SaveSetting "Window3D", "DirectX", "Gravity", CStr(1 / directxGravity)
    SaveSetting "Window3D", "DirectX", "TexFIndex", CStr(directxTexFIndex)
    SaveSetting "Window3D", "DirectX", "Anisotropy", CStr(directxAnisotropy)
    SaveSetting "Window3D", "DirectX", "FovY", CStr(directxFovY)
    SaveSetting "Window3D", "DirectX", "Aspect", CStr(directxAspect)
    Dim batfn As String
    batfn = App.Path & "\" & App.EXEName & ".bat"
    Open batfn For Output As #1
    Print #1, "start " & Replace(LCase(batfn), ".bat", ".exe")
    Close #1
    Shell batfn, vbNormalFocus
    Unload frmMain
    Beep
End Sub

