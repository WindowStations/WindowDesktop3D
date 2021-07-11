VERSION 5.00
Begin VB.Form frmPointer 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Pointer/Point of view"
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
      Height          =   9915
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   12915
      Begin VB.Frame frachkDisable3D 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   6840
         TabIndex        =   45
         Tag             =   "1,20"
         Top             =   1560
         Width           =   5000
         Begin VB.CheckBox chkDisable3D 
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
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   240
            Width           =   200
         End
         Begin VB.Label lblDisable3d 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Disable 3D control"
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
            TabIndex        =   46
            Top             =   240
            Width           =   1920
         End
      End
      Begin VB.Frame frachkDisable2D 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   960
         TabIndex        =   42
         Tag             =   "1,20"
         Top             =   1440
         Width           =   5000
         Begin VB.CheckBox chkDisable2D 
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
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   360
            Width           =   200
         End
         Begin VB.Label lblDisable2D 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Disable 2D control"
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
            TabIndex        =   43
            Top             =   360
            Width           =   1920
         End
      End
      Begin VB.Frame fraSlider8 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   6840
         TabIndex        =   38
         Tag             =   "1,5"
         Top             =   6600
         Width           =   5000
         Begin VB.Label lblValueSlider8 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   15
            TabIndex        =   41
            Top             =   480
            Width           =   15
         End
         Begin VB.Label lblTitleSlider8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Walk acceleration"
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
            TabIndex        =   40
            Top             =   0
            Width           =   1830
         End
         Begin VB.Label lblSlider8 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   0
            TabIndex        =   39
            Top             =   480
            Width           =   5000
         End
      End
      Begin VB.Frame fraSlider7 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   6840
         TabIndex        =   34
         Tag             =   "1,5"
         Top             =   5280
         Width           =   5000
         Begin VB.Label lblValueSlider7 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   15
            TabIndex        =   37
            Top             =   480
            Width           =   15
         End
         Begin VB.Label lblTitleSlider7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Walk speed"
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
            TabIndex        =   36
            Top             =   0
            Width           =   1200
         End
         Begin VB.Label lblSlider7 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   0
            TabIndex        =   35
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
         TabIndex        =   30
         Tag             =   "1,50"
         Top             =   3960
         Width           =   5000
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
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   15
            TabIndex        =   33
            Top             =   480
            Width           =   15
         End
         Begin VB.Label lblTitleSlider6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Point of view acceleration"
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
            Top             =   0
            Width           =   2655
         End
         Begin VB.Label lblSlider6 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   0
            TabIndex        =   31
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
         TabIndex        =   26
         Tag             =   "1,50"
         Top             =   2640
         Width           =   5000
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
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   15
            TabIndex        =   29
            Top             =   480
            Width           =   15
         End
         Begin VB.Label lblTitleSlider5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Point of view speed"
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
            TabIndex        =   28
            Top             =   0
            Width           =   2025
         End
         Begin VB.Label lblSlider5 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   0
            TabIndex        =   27
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
         TabIndex        =   22
         Tag             =   "1,50"
         Top             =   6600
         Width           =   5000
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
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   15
            TabIndex        =   25
            Top             =   480
            Width           =   15
         End
         Begin VB.Label lblTitleSlider4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Wheel acceleration"
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
            TabIndex        =   24
            Top             =   0
            Width           =   1965
         End
         Begin VB.Label lblSlider4 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   0
            TabIndex        =   23
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
         TabIndex        =   18
         Tag             =   "1,50"
         Top             =   5280
         Width           =   5000
         Begin VB.Label lblTitleSlider3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Wheel speed"
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
            TabIndex        =   20
            Top             =   0
            Width           =   1335
         End
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
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   15
            TabIndex        =   19
            Top             =   480
            Width           =   15
         End
         Begin VB.Label lblSlider3 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   0
            TabIndex        =   21
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
         TabIndex        =   14
         Tag             =   "1,50"
         Top             =   4080
         Width           =   5000
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
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   15
            TabIndex        =   17
            Top             =   480
            Width           =   15
         End
         Begin VB.Label lblTitleSlider2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pointer acceleration"
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
            TabIndex        =   16
            Top             =   0
            Width           =   2055
         End
         Begin VB.Label lblSlider2 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   0
            TabIndex        =   15
            Top             =   480
            Width           =   5000
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2760
         TabIndex        =   11
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
            TabIndex        =   12
            Top             =   135
            Width           =   750
         End
      End
      Begin VB.Frame fraSlider1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   960
         TabIndex        =   8
         Tag             =   "1,20"
         Top             =   2640
         Width           =   5000
         Begin VB.Label lblTitleSlider1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pointer speed"
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
            Width           =   1425
         End
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
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   15
            TabIndex        =   9
            Top             =   480
            Width           =   15
         End
         Begin VB.Label lblSlider1 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   0
            TabIndex        =   13
            Top             =   480
            Width           =   5000
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   10200
         TabIndex        =   6
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
            TabIndex        =   7
            Top             =   135
            Width           =   570
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   8280
         TabIndex        =   4
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
            TabIndex        =   5
            Top             =   135
            Width           =   600
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   840
         TabIndex        =   2
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
            TabIndex        =   3
            Top             =   135
            Width           =   480
         End
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   " 2D Pointer                           3D Point of view"
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
         Width           =   11145
      End
   End
End
Attribute VB_Name = "frmPointer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private isloaded As Boolean
Private dragx As Long
Private dragy As Long
Private mdown As Boolean
Private Sub Form_Load()
   On Error Resume Next
   CreateRoundRectFromWindow Me
   CreateRoundRectFromWindow Frame1
   chkDisable2D.Value = pointerDisable2D
   chkDisable3D.Value = pointerDisable3D
   lblTitleSlider1.Tag = pointerMaxPointerSpeed
   lblTitleSlider1.Caption = "Pointer speed " & pointerMaxPointerSpeed
   lblValueSlider1.Width = fraSlider1.Width * (pointerMaxPointerSpeed / 20)
   lblTitleSlider2.Tag = pointerMaxPointerAcceleration
   lblTitleSlider2.Caption = "Pointer acceleration " & pointerMaxPointerAcceleration
   lblValueSlider2.Width = fraSlider2.Width * (pointerMaxPointerAcceleration / 50)
   lblTitleSlider3.Tag = pointerMaxWheelSpeed
   lblTitleSlider3.Caption = "Wheel speed " & pointerMaxWheelSpeed
   lblValueSlider3.Width = fraSlider3.Width * (pointerMaxWheelSpeed / 50)
   lblTitleSlider4.Tag = pointerMaxWheelAcceleration
   lblTitleSlider4.Caption = "Wheel acceleration " & pointerMaxWheelAcceleration
   lblValueSlider4.Width = fraSlider4.Width * (pointerMaxWheelAcceleration / 50)
   lblTitleSlider5.Tag = pointerMaxPOVSpeed
   lblTitleSlider5.Caption = "Point of view speed " & pointerMaxPOVSpeed
   lblValueSlider5.Width = fraSlider5.Width * (pointerMaxPOVSpeed / 50)
   lblTitleSlider6.Tag = pointerMaxPOVAcceleration
   lblTitleSlider6.Caption = "Point of view acceleration " & pointerMaxPOVAcceleration
   lblValueSlider6.Width = fraSlider6.Width * (pointerMaxPOVAcceleration / 50)
   lblTitleSlider7.Tag = pointerMaxWalkSpeed
   lblTitleSlider7.Caption = "Walk speed " & pointerMaxWalkSpeed
   lblValueSlider7.Width = fraSlider7.Width * (pointerMaxWalkSpeed / 5)
   lblTitleSlider8.Tag = pointerMaxWalkAcceleration
   lblTitleSlider8.Caption = "Walk acceleration " & pointerMaxWalkAcceleration
   lblValueSlider8.Width = fraSlider8.Width * (pointerMaxWalkAcceleration / 5)
   dragx = -1
   dragy = -1
   isloaded = True
End Sub
Private Sub Form_Activate()
   WindowTransparency Me.hWnd, displayTransSettings, vbBlack
   lblTitle.TabIndex = 0
End Sub
Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   dragx = x
   dragy = y
End Sub
Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   dragx = -1
   dragy = -1
End Sub
Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Frame1_MouseDown Button, Shift, x, y
End Sub
Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Frame1_MouseMove Button, Shift, x, y
End Sub
Private Sub lblTitle_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   Frame1_MouseUp Button, Shift, x, y
End Sub
Private Sub Frame2_Click()
   frmSettings.show
   frmSettings.top = Me.top
   frmSettings.left = Me.left
   frmMain.SetWindowPos frmSettings.hWnd, -1, 0, 0, 0, 0, False, False
   Unload Me
End Sub
Private Sub chkDisable2D_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   CheckBoxSetting1
End Sub
Private Sub lblDisable2D_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If chkDisable2D.Value = vbUnchecked Then
      chkDisable2D.Value = vbChecked
   Else
      chkDisable2D.Value = vbUnchecked
   End If
   CheckBoxSetting1
End Sub
Private Sub CheckBoxSetting1()
   pointerDisable2D = chkDisable2D.Value
End Sub
Private Sub chkDisable3D_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   CheckBoxSetting2
End Sub
Private Sub lblDisable3d_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If chkDisable3D.Value = vbUnchecked Then
      chkDisable3D.Value = vbChecked
   Else
      chkDisable3D.Value = vbUnchecked
   End If
   CheckBoxSetting2
End Sub
Private Sub CheckBoxSetting2()
   pointerDisable3D = chkDisable3D.Value
End Sub
Private Sub Frame3_Click()
   On Error Resume Next
   pointerMaxPointerSpeed = lblTitleSlider1.Tag
   pointerMaxPointerAcceleration = lblTitleSlider2.Tag ' traPointerAcceleration.Value
   pointerMaxWheelSpeed = lblTitleSlider3.Tag 'traWheelSpeed.Value
   pointerMaxWheelAcceleration = lblTitleSlider4.Tag ' traWheelAcceleration.Value
   pointerMaxPOVSpeed = lblTitleSlider5.Tag 'traPOVSpeed.Value
   pointerMaxPOVAcceleration = lblTitleSlider6.Tag 'traPOVAcceleration.Value
   pointerMaxWalkSpeed = lblTitleSlider7.Tag 'traWalkSpeed.Value
   pointerMaxWalkAcceleration = lblTitleSlider8.Tag 'traWalkAcceleration.Value
   pointerDisable2D = chkDisable2D.Value
   pointerDisable3D = chkDisable3D.Value
   frmMain.Visible = Not CBool(pointerDisable3D)
   RenderEnabled = Not CBool(pointerDisable3D)
   SaveSetting "Window3D", "Pointer", "MaxPointerSpeed", CStr(pointerMaxPointerSpeed)
   SaveSetting "Window3D", "Pointer", "MaxPOVSpeed", CStr(pointerMaxPOVSpeed)
   SaveSetting "Window3D", "Pointer", "MaxPointerAcceleration", CStr(pointerMaxPointerAcceleration)
   SaveSetting "Window3D", "Pointer", "MaxPOVAcceleration", CStr(pointerMaxPOVAcceleration)
   SaveSetting "Window3D", "Pointer", "MaxWheelSpeed", CStr(pointerMaxWheelSpeed)
   SaveSetting "Window3D", "Pointer", "MaxWalkSpeed", CStr(pointerMaxWalkSpeed)
   SaveSetting "Window3D", "Pointer", "MaxWheelAcceleration", CStr(pointerMaxWheelAcceleration)
   SaveSetting "Window3D", "Pointer", "MaxWalkAcceleration", CStr(pointerMaxWalkAcceleration)
   SaveSetting "Window3D", "Pointer", "Disable2D", CStr(pointerDisable2D)
   SaveSetting "Window3D", "Pointer", "Disable3D", CStr(pointerDisable3D)
   Beep
End Sub
Private Sub Frame4_Click()
   Unload Me
End Sub
Private Sub Frame6_Click()
   pointerMaxPointerSpeed = 6
   pointerMaxPOVSpeed = 10
   pointerMaxPointerAcceleration = 20
   pointerMaxPOVAcceleration = 40
   pointerMaxWheelSpeed = 10
   pointerMaxWalkSpeed = 1
   pointerMaxWheelAcceleration = 4
   pointerMaxWalkAcceleration = 5
   pointerDisable2D = 0
   pointerDisable3D = 0
End Sub
Private Sub lblDefault_Click()
   Frame6_Click
End Sub
'
'
'
Private Function UpdateSlider(ByVal x As Single, ByRef sldr As Frame, ByRef lvl As Label, ByRef lbl As Label, ByVal name As String) As Long
   On Error Resume Next
   Dim v As Long
   Dim mm() As String
   Dim min As Long
   Dim max As Long
   mm = Split(sldr.Tag, ",")
   min = mm(0)
   max = mm(1)
   v = (x / sldr.Width) * max
   If v < min Then v = min
   If v > max Then v = max
   If x < 0 Then lvl.Width = 0
   If x >= 0 Then lvl.Width = x
   lbl.Tag = v
   lbl.Caption = name & " " & v
   UpdateSlider = v
End Function
Private Sub lblSlider1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   mdown = True
   pointerMaxPointerSpeed = UpdateSlider(x, fraSlider1, lblValueSlider1, lblTitleSlider1, "Pointer speed")
End Sub
Private Sub lblValueSlider1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider1_MouseDown Button, Shift, x, y
End Sub
Private Sub lblSlider1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If mdown = False Then Exit Sub
   pointerMaxPointerSpeed = UpdateSlider(x, fraSlider1, lblValueSlider1, lblTitleSlider1, "Pointer speed")
End Sub
Private Sub lblValueSlider1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider1_MouseMove Button, Shift, x, y
End Sub
Private Sub lblSlider1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   mdown = False
   pointerMaxPointerSpeed = UpdateSlider(x, fraSlider1, lblValueSlider1, lblTitleSlider1, "Pointer speed")
End Sub
Private Sub lblValueSlider1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider1_MouseUp Button, Shift, x, y
End Sub
''''''''''''''
Private Sub lblSlider2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   mdown = True
   pointerMaxPointerAcceleration = UpdateSlider(x, fraSlider2, lblValueSlider2, lblTitleSlider2, "Pointer acceleration")
End Sub
Private Sub lblValueSlider2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider2_MouseDown Button, Shift, x, y
End Sub
Private Sub lblSlider2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If mdown = False Then Exit Sub
   pointerMaxPointerAcceleration = UpdateSlider(x, fraSlider2, lblValueSlider2, lblTitleSlider2, "Pointer acceleration")
End Sub
Private Sub lblValueSlider2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider2_MouseMove Button, Shift, x, y
End Sub
Private Sub lblSlider2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   mdown = False
   pointerMaxPointerAcceleration = UpdateSlider(x, fraSlider2, lblValueSlider2, lblTitleSlider2, "Pointer acceleration")
End Sub
Private Sub lblValueSlider2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider2_MouseUp Button, Shift, x, y
End Sub
'''''''''''''''''''''''''''''''''
Private Sub lblSlider3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   mdown = True
   pointerMaxWheelSpeed = UpdateSlider(x, fraSlider3, lblValueSlider3, lblTitleSlider3, "Wheel speed")
End Sub
Private Sub lblValueSlider3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider3_MouseDown Button, Shift, x, y
End Sub
Private Sub lblSlider3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If mdown = False Then Exit Sub
   pointerMaxWheelSpeed = UpdateSlider(x, fraSlider3, lblValueSlider3, lblTitleSlider3, "Wheel speed")
End Sub
Private Sub lblValueSlider3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider3_MouseMove Button, Shift, x, y
End Sub
Private Sub lblSlider3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   mdown = False
   pointerMaxWheelSpeed = UpdateSlider(x, fraSlider3, lblValueSlider3, lblTitleSlider3, "Wheel speed")
End Sub
Private Sub lblValueSlider3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider3_MouseUp Button, Shift, x, y
End Sub
'''''''''''''''''''''''''''''''''
Private Sub lblSlider4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   mdown = True
   pointerMaxWheelAcceleration = UpdateSlider(x, fraSlider4, lblValueSlider4, lblTitleSlider4, "Wheel acceleration")
End Sub
Private Sub lblValueSlider4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider4_MouseDown Button, Shift, x, y
End Sub
Private Sub lblSlider4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If mdown = False Then Exit Sub
   pointerMaxWheelAcceleration = UpdateSlider(x, fraSlider4, lblValueSlider4, lblTitleSlider4, "Wheel acceleration")
End Sub
Private Sub lblValueSlider4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider4_MouseMove Button, Shift, x, y
End Sub
Private Sub lblSlider4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   mdown = False
   pointerMaxWheelAcceleration = UpdateSlider(x, fraSlider4, lblValueSlider4, lblTitleSlider4, "Wheel acceleration")
End Sub
Private Sub lblValueSlider4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider4_MouseUp Button, Shift, x, y
End Sub
'''''''''''''''''''''''''''''''''
Private Sub lblSlider5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   mdown = True
   pointerMaxPOVSpeed = UpdateSlider(x, fraSlider5, lblValueSlider5, lblTitleSlider5, "Point of view speed")
End Sub
Private Sub lblValueSlider5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider5_MouseDown Button, Shift, x, y
End Sub
Private Sub lblSlider5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If mdown = False Then Exit Sub
   pointerMaxPOVSpeed = UpdateSlider(x, fraSlider5, lblValueSlider5, lblTitleSlider5, "Point of view speed")
End Sub
Private Sub lblValueSlider5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider5_MouseMove Button, Shift, x, y
End Sub
Private Sub lblSlider5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   mdown = False
   pointerMaxPOVSpeed = UpdateSlider(x, fraSlider5, lblValueSlider5, lblTitleSlider5, "Point of view speed")
End Sub
Private Sub lblValueSlider5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider5_MouseUp Button, Shift, x, y
End Sub
'''''''''''''''''''''''''''''''''
Private Sub lblSlider6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   mdown = True
   pointerMaxPOVAcceleration = UpdateSlider(x, fraSlider6, lblValueSlider6, lblTitleSlider6, "Point of view acceleration")
End Sub
Private Sub lblValueSlider6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider6_MouseDown Button, Shift, x, y
End Sub
Private Sub lblSlider6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If mdown = False Then Exit Sub
   pointerMaxPOVAcceleration = UpdateSlider(x, fraSlider6, lblValueSlider6, lblTitleSlider6, "Point of view acceleration")
End Sub
Private Sub lblValueSlider6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider6_MouseMove Button, Shift, x, y
End Sub
Private Sub lblSlider6_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   mdown = False
   pointerMaxPOVAcceleration = UpdateSlider(x, fraSlider6, lblValueSlider6, lblTitleSlider6, "Point of view acceleration")
End Sub
Private Sub lblValueSlider6_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider6_MouseUp Button, Shift, x, y
End Sub
'''''''''''''''''''''''''''''''''
Private Sub lblSlider7_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   mdown = True
   pointerMaxWalkSpeed = UpdateSlider(x, fraSlider7, lblValueSlider7, lblTitleSlider7, "Walk speed")
End Sub
Private Sub lblValueSlider7_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider7_MouseDown Button, Shift, x, y
End Sub
Private Sub lblSlider7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If mdown = False Then Exit Sub
   pointerMaxWalkSpeed = UpdateSlider(x, fraSlider7, lblValueSlider7, lblTitleSlider7, "Walk speed")
End Sub
Private Sub lblValueSlider7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider7_MouseMove Button, Shift, x, y
End Sub
Private Sub lblSlider7_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   mdown = False
   pointerMaxWalkSpeed = UpdateSlider(x, fraSlider7, lblValueSlider7, lblTitleSlider7, "Walk speed")
End Sub
Private Sub lblValueSlider7_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider7_MouseUp Button, Shift, x, y
End Sub
'''''''''''''''''''''''''''''''''
Private Sub lblSlider8_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   mdown = True
   pointerMaxWalkAcceleration = UpdateSlider(x, fraSlider8, lblValueSlider8, lblTitleSlider8, "Walk acceleration")
End Sub
Private Sub lblValueSlider8_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider8_MouseDown Button, Shift, x, y
End Sub
Private Sub lblSlider8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If mdown = False Then Exit Sub
   pointerMaxWalkAcceleration = UpdateSlider(x, fraSlider8, lblValueSlider8, lblTitleSlider8, "Walk acceleration")
End Sub
Private Sub lblValueSlider8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider8_MouseMove Button, Shift, x, y
End Sub
Private Sub lblSlider8_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   mdown = False
   pointerMaxWalkAcceleration = UpdateSlider(x, fraSlider8, lblValueSlider8, lblTitleSlider8, "Walk acceleration")
End Sub
Private Sub lblValueSlider8_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSlider8_MouseUp Button, Shift, x, y
End Sub
Private Sub lblClose_Click()
   Frame4_Click
End Sub
Private Sub lblBack_Click()
   Frame2_Click
End Sub
Private Sub lblApply_Click()
   Frame3_Click
End Sub
Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Frame6.BackColor = &H808080
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
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Frame6.BackColor <> &H404040 Then Frame6.BackColor = &H404040
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
Private Sub UserControl11_GotFocus()
End Sub
