VERSION 5.00
Begin VB.Form frmDisplay 
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
      Begin VB.Frame frachkDisable2D 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   960
         TabIndex        =   32
         Tag             =   "1,20"
         Top             =   1320
         Width           =   5000
         Begin VB.CheckBox chkHideText 
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
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   360
            Width           =   200
         End
         Begin VB.Label lblDisable2D 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hide selected item text"
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
            TabIndex        =   33
            Top             =   360
            Width           =   2340
         End
      End
      Begin VB.Frame fraSlider1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   960
         TabIndex        =   28
         Tag             =   "1,5000"
         Top             =   2640
         Width           =   5000
         Begin VB.Label lblValueSlider1 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   315
            Left            =   15
            TabIndex        =   30
            Top             =   480
            Width           =   15
         End
         Begin VB.Label lblTitleSlider1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Text delay"
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
            TabIndex        =   29
            Top             =   0
            Width           =   5000
         End
         Begin VB.Label lblSlider1 
            AutoSize        =   -1  'True
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
            Height          =   315
            Left            =   0
            TabIndex        =   31
            Top             =   480
            Width           =   5000
         End
      End
      Begin VB.Frame fraSlider2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   960
         TabIndex        =   24
         Tag             =   "1,5"
         Top             =   3960
         Width           =   5000
         Begin VB.Label lblTitleSlider2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Text speed"
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
            TabIndex        =   26
            Top             =   0
            Width           =   5000
         End
         Begin VB.Label lblValueSlider2 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   315
            Left            =   15
            TabIndex        =   25
            Top             =   480
            Width           =   15
         End
         Begin VB.Label lblSlider2 
            AutoSize        =   -1  'True
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
            Height          =   315
            Left            =   0
            TabIndex        =   27
            Top             =   480
            Width           =   5000
         End
      End
      Begin VB.Frame fraSlider3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   960
         TabIndex        =   20
         Tag             =   "1,5000"
         Top             =   5280
         Width           =   5000
         Begin VB.Label lblValueSlider3 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   315
            Left            =   15
            TabIndex        =   22
            Top             =   480
            Width           =   15
         End
         Begin VB.Label lblTitleSlider3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Text fade"
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
            TabIndex        =   21
            Top             =   0
            Width           =   5000
         End
         Begin VB.Label lblSlider3 
            AutoSize        =   -1  'True
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
            Height          =   315
            Left            =   0
            TabIndex        =   23
            Top             =   480
            Width           =   5000
         End
      End
      Begin VB.Frame fraSlider4 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   960
         TabIndex        =   16
         Tag             =   "1,1000"
         Top             =   6600
         Width           =   5000
         Begin VB.Label lblTitleSlider4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Text position"
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
            TabIndex        =   18
            Top             =   0
            Width           =   5000
         End
         Begin VB.Label lblValueSlider4 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   315
            Left            =   15
            TabIndex        =   17
            Top             =   480
            Width           =   15
         End
         Begin VB.Label lblSlider4 
            AutoSize        =   -1  'True
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
            Height          =   315
            Left            =   0
            TabIndex        =   19
            Top             =   480
            Width           =   5000
         End
      End
      Begin VB.Frame fraSlider5 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   6840
         TabIndex        =   12
         Tag             =   "1,255"
         Top             =   1920
         Width           =   5000
         Begin VB.Label lblTitleSlider5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Window translucency 3D"
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
            Top             =   0
            Width           =   5000
         End
         Begin VB.Label lblValueSlider5 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   315
            Left            =   15
            TabIndex        =   13
            Top             =   480
            Width           =   15
         End
         Begin VB.Label lblSlider5 
            AutoSize        =   -1  'True
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
            Height          =   315
            Left            =   0
            TabIndex        =   15
            Top             =   480
            Width           =   5000
         End
      End
      Begin VB.Frame fraSlider6 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   6840
         TabIndex        =   8
         Tag             =   "1,255"
         Top             =   3240
         Width           =   5000
         Begin VB.Label lblTitleSlider6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Window translucency Settings"
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
            TabIndex        =   10
            Top             =   0
            Width           =   5000
         End
         Begin VB.Label lblValueSlider6 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   315
            Left            =   15
            TabIndex        =   9
            Top             =   480
            Width           =   15
         End
         Begin VB.Label lblSlider6 
            AutoSize        =   -1  'True
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
            Height          =   315
            Left            =   0
            TabIndex        =   11
            Top             =   480
            Width           =   5000
         End
      End
      Begin VB.Frame fracmdClose 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   10200
         TabIndex        =   2
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
            TabIndex        =   3
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
         TabIndex        =   4
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
            TabIndex        =   5
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
         TabIndex        =   6
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
            TabIndex        =   7
            Top             =   135
            Width           =   600
         End
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Display"
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
         Top             =   600
         Width           =   11100
      End
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private isloaded  As Boolean
Private dragx As Long
Private dragy As Long
Private mdown As Boolean



Private Sub Form_Load()
    On Error Resume Next
    CreateRoundRectFromWindow Me
    CreateRoundRectFromWindow fraMain
    chkHideText.Value = displayHide
    
    lblTitleSlider1.Tag = displayDelay
    lblTitleSlider1.Caption = "Text delay " & CStr(displayDelay)
    lblValueSlider1.Width = fraSlider1.Width * (displayDelay / 5000)
    lblTitleSlider2.Tag = displaySpeed
    lblTitleSlider2.Caption = "Text speed " & CStr(displaySpeed)
    lblValueSlider2.Width = fraSlider2.Width * (displaySpeed / 5)
    lblTitleSlider3.Tag = displayFade
    lblTitleSlider3.Caption = "Text fade " & CStr(displayFade)
    lblValueSlider3.Width = fraSlider3.Width * (displayFade / 5000)
    lblTitleSlider4.Tag = displayPosition '- 1000
    lblTitleSlider4.Caption = "Text position " & CStr(displayPosition)
    lblValueSlider4.Width = fraSlider4.Width * (displayPosition / 2000)
    lblTitleSlider5.Tag = displayTrans3D
    lblTitleSlider5.Caption = "Window translucency 3D " & CStr(displayTrans3D)
    lblValueSlider5.Width = fraSlider5.Width * (displayTrans3D / 255)
    lblTitleSlider6.Tag = displayTransSettings
    lblTitleSlider6.Caption = "Window translucency Settings " & CStr(displayTransSettings)
    lblValueSlider6.Width = fraSlider6.Width * (displayTransSettings / 255)
    dragx = -1
    dragy = -1
    isloaded = True
End Sub
Private Sub Form_Activate()
    WindowTransparency Me.hWnd, displayTransSettings, vbBlack
End Sub
Private Sub fraMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    dragx = X
    dragy = Y
End Sub
Private Sub fraMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    dragx = -1
    dragy = -1
End Sub
Private Sub fraTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraMain_MouseDown Button, Shift, X, Y
End Sub


Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraMain_MouseDown Button, Shift, X, Y
End Sub
Private Sub fraTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraMain_MouseMove Button, Shift, X, Y
End Sub
Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraMain_MouseMove Button, Shift, X, Y
End Sub
Private Sub fraTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraMain_MouseUp Button, Shift, X, Y
End Sub
Private Sub lblTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraMain_MouseUp Button, Shift, X, Y
End Sub
Private Sub fraMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If fracmdBack.BackColor <> &H404040 Then fracmdBack.BackColor = &H404040
    If fracmdApply.BackColor <> &H404040 Then fracmdApply.BackColor = &H404040
    If fracmdClose.BackColor <> &H404040 Then fracmdClose.BackColor = &H404040
    If dragx > -1 Then
        If X > dragx Then
            Me.left = Me.left + (X - dragx)
        ElseIf X < dragx Then
            Me.left = Me.left - (dragx - X)
        End If
    End If
    If dragy > -1 Then
        If Y > dragy Then
            Me.top = Me.top + (Y - dragy)
        ElseIf Y < dragy Then
            Me.top = Me.top - (dragy - Y)
        End If
    End If
End Sub
Private Sub fracmdBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fracmdBack.BackColor = &H808080
End Sub
Private Sub fracmdApply_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fracmdApply.BackColor = &H808080
End Sub
Private Sub fracmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fracmdClose.BackColor = &H808080
End Sub
Private Sub fracmdBack_Click()
    frmSettings.show
    frmSettings.top = Me.top
    frmSettings.left = Me.left
    frmMain.SetWindowPos frmSettings.hWnd, -1, 0, 0, 0, 0, False, False
    Unload Me
End Sub
Private Sub lblBack_Click()
    fracmdBack_Click
End Sub


Private Sub chkHideText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  CheckBoxSetting
End Sub


Private Sub lblDisable2D_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If chkHideText.Value = vbUnchecked Then
        chkHideText.Value = vbChecked
    Else
        chkHideText.Value = vbUnchecked
    End If
    CheckBoxSetting
End Sub
Private Sub CheckBoxSetting()
   displayHide = chkHideText.Value
End Sub
Private Sub fracmdApply_Click()
    On Error Resume Next
    displayDelay = CLng(lblTitleSlider1.Tag)
    displaySpeed = CLng(lblTitleSlider2.Tag)
    displayFade = CLng(lblTitleSlider3.Tag)
    displayPosition = CLng(lblTitleSlider4.Tag)
    displayTrans3D = CInt(lblTitleSlider5.Tag)
    displayTransSettings = CInt(lblTitleSlider6.Tag)
    displayHide = chkHideText.Value
    
    SaveSetting "Window3D", "Display", "Delay", CStr(displayDelay)
    SaveSetting "Window3D", "Display", "Speed", CStr(displaySpeed)
    SaveSetting "Window3D", "Display", "Fade", CStr(displayFade)
    SaveSetting "Window3D", "Display", "Position", CStr(displayPosition)
    SaveSetting "Window3D", "Display", "3D", CStr(displayTrans3D)
    SaveSetting "Window3D", "Display", "Settings", CStr(displayTransSettings)
    SaveSetting "Window3D", "Display", "Hide", CStr(displayHide)
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
Private Function UpdateSlider(ByVal X As Single, ByRef sldr As Frame, ByRef lvl As Label, ByRef lbl As Label, ByVal name As String) As Long
    On Error Resume Next
    Dim v As Long
    Dim mm() As String
    Dim min As Long
    Dim max As Long
    mm = Split(sldr.Tag, ",")
    min = mm(0)
    max = mm(1)
    v = (X / sldr.Width) * max
    If v < min Then v = min
    If v > max Then v = max
    If X < 0 Then lvl.Width = 0
    If X >= 0 Then lvl.Width = X
    lbl.Tag = v
    lbl.Caption = name & " " & v
    UpdateSlider = v
End Function

Private Sub lblSlider1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mdown = True
    displayDelay = UpdateSlider(X, fraSlider1, lblValueSlider1, lblTitleSlider1, "Text delay")
End Sub
Private Sub lblValueSlider1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSlider1_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblSlider1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mdown = False Then Exit Sub
    displayDelay = UpdateSlider(X, fraSlider1, lblValueSlider1, lblTitleSlider1, "Text delay")
End Sub
Private Sub lblValueSlider1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSlider1_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblSlider1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mdown = False
End Sub
Private Sub lblValueSlider1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSlider1_MouseUp Button, Shift, X, Y
End Sub
''''''''''''''

Private Sub lblSlider2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mdown = True
    displaySpeed = UpdateSlider(X, fraSlider2, lblValueSlider2, lblTitleSlider2, "Text speed")
End Sub
Private Sub lblValueSlider2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSlider2_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblSlider2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mdown = False Then Exit Sub
    displaySpeed = UpdateSlider(X, fraSlider2, lblValueSlider2, lblTitleSlider2, "Text speed")
End Sub
Private Sub lblValueSlider2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSlider2_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblSlider2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mdown = False
End Sub
Private Sub lblValueSlider2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSlider2_MouseUp Button, Shift, X, Y
End Sub
'''''''''''''''''''''''''''''''''

Private Sub lblSlider3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mdown = True
    displayFade = UpdateSlider(X, fraSlider3, lblValueSlider3, lblTitleSlider3, "Text fade")
End Sub
Private Sub lblValueSlider3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSlider3_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblSlider3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mdown = False Then Exit Sub
    displayFade = UpdateSlider(X, fraSlider3, lblValueSlider3, lblTitleSlider3, "Text fade")
End Sub
Private Sub lblValueSlider3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSlider3_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblSlider3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mdown = False
End Sub
Private Sub lblValueSlider3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSlider3_MouseUp Button, Shift, X, Y
End Sub
'''''''''''''''''''''''''''''''''

Private Sub lblSlider4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mdown = True
    displayPosition = UpdateSlider(X, fraSlider4, lblValueSlider4, lblTitleSlider4, "Text position") '- 1000
End Sub
Private Sub lblValueSlider4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSlider4_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblSlider4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mdown = False Then Exit Sub
    displayPosition = UpdateSlider(X, fraSlider4, lblValueSlider4, lblTitleSlider4, "Text position") ' - 1000
End Sub
Private Sub lblValueSlider4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSlider4_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblSlider4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mdown = False
End Sub
Private Sub lblValueSlider4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSlider4_MouseUp Button, Shift, X, Y
End Sub
'''''''''''''''''''''''''''''''''

Private Sub lblSlider5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mdown = True
    displayTrans3D = CInt(UpdateSlider(X, fraSlider5, lblValueSlider5, lblTitleSlider5, "Window translucency 3D"))
    WindowTransparency frmMain.hWnd, displayTrans3D, vbBlack
End Sub
Private Sub lblValueSlider5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSlider5_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblSlider5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mdown = False Then Exit Sub
    displayTrans3D = CInt(UpdateSlider(X, fraSlider5, lblValueSlider5, lblTitleSlider5, "Window translucency 3D"))
    WindowTransparency frmMain.hWnd, displayTrans3D, vbBlack
End Sub
Private Sub lblValueSlider5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSlider5_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblSlider5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mdown = False
End Sub
Private Sub lblValueSlider5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSlider5_MouseUp Button, Shift, X, Y
End Sub
'''''''''''''''''''''''''''''''''

Private Sub lblSlider6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mdown = True
    displayTransSettings = CInt(UpdateSlider(X, fraSlider6, lblValueSlider6, lblTitleSlider6, "Window translucency Settings"))
    WindowTransparency Me.hWnd, displayTransSettings, vbBlack
End Sub
Private Sub lblValueSlider6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSlider6_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblSlider6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mdown = False Then Exit Sub
    displayTransSettings = CInt(UpdateSlider(X, fraSlider6, lblValueSlider6, lblTitleSlider6, "Window translucency Settings"))
    WindowTransparency Me.hWnd, displayTransSettings, vbBlack
End Sub
Private Sub lblValueSlider6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSlider6_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblSlider6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mdown = False
End Sub
Private Sub lblValueSlider6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSlider6_MouseUp Button, Shift, X, Y
End Sub
'''''''''''''''''''''''''''''''''
'
'
'Private Sub lblSlider7_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    mdown = True
'   pointerMaxWalkSpeed = UpdateSlider(x, fraSlider7, lblValueSlider7, lblTitleSlider7, "Walk speed")
'End Sub
'Private Sub lblValueSlider7_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    lblSlider7_MouseDown Button, Shift, x, y
'End Sub
'
'Private Sub lblSlider7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If mdown = False Then Exit Sub
'      pointerMaxWalkSpeed = UpdateSlider(x, fraSlider7, lblValueSlider7, lblTitleSlider7, "Walk speed")
'End Sub
'Private Sub lblValueSlider7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'   lblSlider7_MouseMove Button, Shift, x, y
'End Sub
'
'Private Sub lblSlider7_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'   mdown = False
'     pointerMaxWalkSpeed = UpdateSlider(x, fraSlider7, lblValueSlider7, lblTitleSlider7, "Walk speed")
'End Sub
'Private Sub lblValueSlider7_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'   lblSlider7_MouseUp Button, Shift, x, y
'End Sub
''''''''''''''''''''''''''''''''''
'
'
'
'Private Sub lblSlider8_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    mdown = True
'     pointerMaxWalkAcceleration = UpdateSlider(x, fraSlider8, lblValueSlider8, lblTitleSlider8, "Walk acceleration")
'End Sub
'Private Sub lblValueSlider8_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    lblSlider8_MouseDown Button, Shift, x, y
'End Sub
'
'Private Sub lblSlider8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If mdown = False Then Exit Sub
'     pointerMaxWalkAcceleration = UpdateSlider(x, fraSlider8, lblValueSlider8, lblTitleSlider8, "Walk acceleration")
'End Sub
'Private Sub lblValueSlider8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'   lblSlider8_MouseMove Button, Shift, x, y
'End Sub
'
'Private Sub lblSlider8_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'   mdown = False
'    pointerMaxWalkAcceleration = UpdateSlider(x, fraSlider8, lblValueSlider8, lblTitleSlider8, "Walk acceleration")
'End Sub
'Private Sub lblValueSlider8_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'   lblSlider8_MouseUp Button, Shift, x, y
'End Sub
