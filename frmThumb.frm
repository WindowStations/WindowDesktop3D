VERSION 5.00
Begin VB.Form frmThumb 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Thumb stick calibration"
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
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   7080
      Top             =   8640
   End
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
      Height          =   9915
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   12915
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4680
         TabIndex        =   28
         Top             =   8400
         Width           =   1695
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "R Default"
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
            TabIndex        =   29
            Top             =   135
            Width           =   960
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2760
         TabIndex        =   26
         Top             =   8400
         Width           =   1695
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "L Default"
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
            TabIndex        =   27
            Top             =   135
            Width           =   930
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5895
         Left            =   6720
         TabIndex        =   9
         Top             =   1800
         Width           =   5175
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   2000
            Left            =   1080
            ScaleHeight     =   133
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   133
            TabIndex        =   33
            Top             =   1080
            Width           =   2000
         End
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
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
            Height          =   495
            Left            =   720
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "4000"
            Top             =   4560
            Width           =   1080
         End
         Begin VB.TextBox Text7 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
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
            Height          =   495
            Left            =   720
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   "4000"
            Top             =   3840
            Width           =   1095
         End
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
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
            Height          =   495
            Left            =   2880
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "4000"
            Top             =   4560
            Width           =   1080
         End
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
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
            Height          =   495
            Left            =   2880
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   "4000"
            Top             =   3840
            Width           =   1080
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Right"
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
            Left            =   30
            TabIndex        =   41
            Top             =   4560
            Width           =   555
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Left"
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
            TabIndex        =   40
            Top             =   3840
            Width           =   390
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Down"
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
            Left            =   2160
            TabIndex        =   39
            Top             =   4560
            Width           =   615
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Up"
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
            Left            =   2520
            TabIndex        =   38
            Top             =   3840
            Width           =   300
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Calibrate right thumb stick dead zone"
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
            Left            =   120
            TabIndex        =   30
            Top             =   0
            Width           =   3885
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "y"
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
            Left            =   2040
            TabIndex        =   25
            Top             =   3120
            Width           =   120
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "x"
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
            Left            =   3240
            TabIndex        =   24
            Top             =   1920
            Width           =   105
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "y"
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
            Left            =   2040
            TabIndex        =   23
            Top             =   600
            Width           =   120
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "x"
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
            Left            =   795
            TabIndex        =   22
            Top             =   1920
            Width           =   105
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5895
         Left            =   840
         TabIndex        =   8
         Top             =   1800
         Width           =   5175
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            FillColor       =   &H00FFFFFF&
            ForeColor       =   &H00FFFFFF&
            Height          =   2000
            Left            =   1080
            ScaleHeight     =   133
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   133
            TabIndex        =   32
            Top             =   1080
            Width           =   2000
         End
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
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
            Height          =   495
            Left            =   720
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "4000"
            Top             =   4560
            Width           =   1080
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
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
            Height          =   495
            Left            =   720
            TabIndex        =   12
            TabStop         =   0   'False
            Text            =   "4000"
            Top             =   3840
            Width           =   1095
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
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
            Height          =   495
            Left            =   2880
            TabIndex        =   11
            TabStop         =   0   'False
            Text            =   "4000"
            Top             =   4560
            Width           =   1080
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
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
            Height          =   495
            Left            =   2880
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "4000"
            Top             =   3840
            Width           =   1095
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Right"
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
            Left            =   30
            TabIndex        =   37
            Top             =   4560
            Width           =   555
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Left"
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
            TabIndex        =   36
            Top             =   3840
            Width           =   390
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Down"
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
            Left            =   2160
            TabIndex        =   35
            Top             =   4560
            Width           =   615
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Up"
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
            Left            =   2520
            TabIndex        =   34
            Top             =   3840
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Calibrate left thumb stick dead zone"
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
            Left            =   120
            TabIndex        =   31
            Top             =   0
            Width           =   3720
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "y"
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
            Left            =   2040
            TabIndex        =   17
            Top             =   3120
            Width           =   120
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "x"
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
            Left            =   3240
            TabIndex        =   16
            Top             =   1920
            Width           =   105
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "y"
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
            Left            =   2040
            TabIndex        =   15
            Top             =   600
            Width           =   120
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "x"
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
            Left            =   795
            TabIndex        =   14
            Top             =   1920
            Width           =   105
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   8280
         TabIndex        =   6
         Top             =   8400
         Width           =   1695
         Begin VB.Label Label14 
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
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   10200
         TabIndex        =   4
         Top             =   8400
         Width           =   1695
         Begin VB.Label Label13 
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
            TabIndex        =   5
            Top             =   135
            Width           =   570
         End
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   840
         TabIndex        =   2
         Top             =   8400
         Width           =   1695
         Begin VB.Label Label12 
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
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Thumb stick calibration"
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
Attribute VB_Name = "frmThumb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private isloaded As Boolean
Private dragx As Long
Private dragy As Long
'Private Declare Function Arc Lib "gdi32" (ByVal hdc As Long, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long, ByVal nXStartArc As Long, ByVal nYStartArc As Long, ByVal nXEndArc As Long, ByVal nYEndArc As Long) As Long
'Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long) As Long
'Private Const Pi As Double = 3.14159265358979
' When an arc or a partial circle or ellipse is drawn, StartArc and EndArc specify (in radians) the beginning and end positions of the arc.
' The range for both is 2 pi radians to 2 pi radians. The default value for StartArc is 0 radians; the default for EndArc is 2 * pi radians.
Sub DrawCircle(p As PictureBox, x As Single, y As Single, Radius As Single, Optional Aspect As Single = 1, Optional StartArc, Optional EndArc)
   Dim iXStartArc As Long, iYStartArc As Long, iXEndArc As Long, iYEndArc As Long
   Dim iAspectX As Single
   Dim iAspectY As Single
   Dim iStartArc As Single
   Dim iEndArc As Single
   Dim iDontDraw As Boolean
   Dim iSolidFigure As Boolean
   p.Cls
   ' VB Circle
   If IsMissing(StartArc) And IsMissing(EndArc) Then
      p.Circle (x, y), Radius, , , , Aspect
      If p.FillStyle = vbSolid Then
         iSolidFigure = True
      End If
   ElseIf IsMissing(StartArc) Then
      p.Circle (x, y), Radius, , , EndArc, Aspect
   ElseIf IsMissing(EndArc) Then
      p.Circle (x, y), Radius, , StartArc, , Aspect
   Else
      p.Circle (x, y), Radius, , StartArc, EndArc, Aspect
   End If
   '    ' API
   '    If Aspect > 1 Then
   '        iAspectX = 1 / Aspect
   '        iAspectY = 1
   '    Else
   '        iAspectX = 1
   '        iAspectY = 1 * Aspect
   '    End If
   '
   '    If IsMissing(StartArc) Then
   '        iStartArc = 0
   '    Else
   '        iStartArc = StartArc
   '    End If
   '    If IsMissing(EndArc) Then
   '        iEndArc = 0 '2 * Pi
   '    Else
   '        iEndArc = EndArc
   '    End If
   '
   '    If Not IsMissing(EndArc) Then
   '        If iStartArc = 0 And iEndArc >= 6.2768 Then iDontDraw = True
   '    End If
   '
   '    If Not iDontDraw Then
   '        iXStartArc = Radius * iAspectX * Cos(iStartArc) + x
   '        iYStartArc = Radius * iAspectY * Sin(iStartArc) * -1 + y
   '        iXEndArc = Radius * iAspectX * Cos(iEndArc) + x
   '        iYEndArc = Radius * iAspectY * Sin(iEndArc) * -1 + y
   '
   '        If iSolidFigure Then
   '            Ellipse Picture2.hdc, x - Radius * iAspectX, y - Radius * iAspectY, x + Radius * iAspectX, y + Radius * iAspectY
   '        Else
   '            Arc Picture2.hdc, x - Radius * iAspectX, y - Radius * iAspectY, x + Radius * iAspectX, y + Radius * iAspectY, iXStartArc, iYStartArc, iXEndArc, iYEndArc
   '        End If
   '        Picture2.Refresh
   '    End If
   '
   '    Picture2.CurrentX = x
   '    Picture2.CurrentY = y
End Sub
Private Sub Form_Load()
   On Error Resume Next
   CreateRoundRectFromWindow Me
   CreateRoundRectFromWindow Frame1
   '   Picture1.AutoRedraw = True
   '   Picture1.Line (0, 0)-(500, 500), QBColor(0)
   '   Picture1.Line (500, 0)-(1000, 500), QBColor(1), B
   '   Picture1.Line (1000, 0)-(1500, 500), QBColor(2), BF
   '   Picture1.Circle (Picture1.Width / 2, Picture1.Height / 2), QBColor(3), 2000
   CreateRoundRectFromWindow2 Picture1
   CreateRoundRectFromWindow2 Picture2
   Text1.Text = CStr(keymapLThumbUp)
   Text2.Text = CStr(keymapLThumbDown)
   Text3.Text = CStr(keymapLThumbLeft)
   Text4.Text = CStr(keymapLThumbRight)
   Text5.Text = CStr(keymapRThumbUp)
   Text6.Text = CStr(keymapRThumbDown)
   Text7.Text = CStr(keymapRThumbLeft)
   Text8.Text = CStr(keymapRThumbRight)
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
Private Sub Text1_Change()
   If isloaded = False Then Exit Sub
   TextxChange Text1
End Sub
Private Sub Text2_Change()
   If isloaded = False Then Exit Sub
   TextxChange Text2
End Sub
Private Sub Text3_Change()
   If isloaded = False Then Exit Sub
   TextxChange Text3
End Sub
Private Sub Text4_Change()
   If isloaded = False Then Exit Sub
   TextxChange Text4
End Sub
Private Sub Text5_Change()
   If isloaded = False Then Exit Sub
   TextxChange Text5
End Sub
Private Sub Text6_Change()
   If isloaded = False Then Exit Sub
   TextxChange Text6
End Sub
Private Sub Text7_Change()
   If isloaded = False Then Exit Sub
   TextxChange Text7
End Sub
Private Sub Text8_Change()
   If isloaded = False Then Exit Sub
   TextxChange Text8
End Sub
Private Sub Frame8_Click()
   frmSettings.show
   frmSettings.top = Me.top
   frmSettings.left = Me.left
   frmMain.SetWindowPos frmSettings.hWnd, -1, 0, 0, 0, 0, False, False
   Unload Me
End Sub
Private Sub Label12_Click()
   Frame8_Click
End Sub
Private Sub Frame6_Click()
   keymapLThumbUp = 4000
   keymapLThumbDown = -4000
   keymapLThumbLeft = -4000
   keymapLThumbRight = 4000
   Text1.Text = CStr(keymapLThumbUp): Text1.Tag = Text1.Text
   Text2.Text = CStr(keymapLThumbDown): Text2.Tag = Text2.Text
   Text3.Text = CStr(keymapLThumbLeft): Text3.Tag = Text3.Text
   Text4.Text = CStr(keymapLThumbRight): Text4.Tag = Text4.Text
   Beep
End Sub
Private Sub Label15_Click()
   Frame6_Click
End Sub
Private Sub Frame7_Click()
   keymapRThumbUp = 4000
   keymapRThumbDown = -4000
   keymapRThumbLeft = -4000
   keymapRThumbRight = 4000
   Text5.Text = CStr(keymapRThumbUp): Text5.Tag = Text5.Text
   Text6.Text = CStr(keymapRThumbDown): Text6.Tag = Text6.Text
   Text7.Text = CStr(keymapRThumbLeft): Text7.Tag = Text7.Text
   Text8.Text = CStr(keymapRThumbRight): Text8.Tag = Text8.Text
   Beep
End Sub
Private Sub Label16_Click()
   Frame7_Click
End Sub
Private Sub Frame3_Click()
   On Error Resume Next
   ValidateDeadzone
   keymapLThumbUp = CLng(Text1.Text)  ' CLng(GetSetting("Window3D", "ButtonMap", "LThumbUp", "5000"))   ' left thumb dead zone
   keymapLThumbDown = CLng(Text2.Text)  ' CLng(GetSetting("Window3D", "ButtonMap", "LThumbDown", "5000"))
   keymapLThumbLeft = CLng(Text3.Text)  ' CLng(GetSetting("Window3D", "ButtonMap", "LThumbLeft", "5000"))
   keymapLThumbRight = CLng(Text4.Text)  '  CLng(GetSetting("Window3D", "ButtonMap", "LThumbRight", "5000"))
   keymapRThumbUp = CLng(Text5.Text)  '  CLng(GetSetting("Window3D", "ButtonMap", "RThumbUp", "5000"))   ' right thumb dead zone
   keymapRThumbDown = CLng(Text6.Text)  ' CLng(GetSetting("Window3D", "ButtonMap", "RThumbDown", "5000"))
   keymapRThumbLeft = CLng(Text7.Text)  ' CLng(GetSetting("Window3D", "ButtonMap", "RThumbLeft", "5000"))
   keymapRThumbRight = CLng(Text8.Text)  '  CLng(GetSetting("Window3D", "ButtonMap", "RThumbRight", "5000"))
   SaveSetting "Window3D", "ButtonMap", "LThumbUp", keymapLThumbUp
   SaveSetting "Window3D", "ButtonMap", "LThumbDown", keymapLThumbDown
   SaveSetting "Window3D", "ButtonMap", "LThumbLeft", keymapLThumbLeft
   SaveSetting "Window3D", "ButtonMap", "LThumbRight", keymapLThumbRight
   SaveSetting "Window3D", "ButtonMap", "RThumbUp", keymapRThumbUp
   SaveSetting "Window3D", "ButtonMap", "RThumbDown", keymapRThumbDown
   SaveSetting "Window3D", "ButtonMap", "RThumbLeft", keymapRThumbLeft
   SaveSetting "Window3D", "ButtonMap", "RThumbRight", keymapRThumbRight
   Beep
End Sub
Private Sub Label14_Click()
   Frame3_Click
End Sub
Private Sub Label13_Click()
   Frame2_Click
End Sub
Private Sub Frame2_Click()
   Unload Me
End Sub
Private Sub Frame8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Frame8.BackColor = &H808080
End Sub
Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Frame6.BackColor = &H808080
End Sub
Private Sub Frame7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Frame7.BackColor = &H808080
End Sub
Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Frame3.BackColor = &H808080
End Sub
Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Frame2.BackColor = &H808080
End Sub
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Frame8.BackColor <> &H404040 Then Frame8.BackColor = &H404040
   If Frame6.BackColor <> &H404040 Then Frame6.BackColor = &H404040
   If Frame7.BackColor <> &H404040 Then Frame7.BackColor = &H404040
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
Private Sub ValidateDeadzone()
   isloaded = False
   If IsNumeric(Text1.Text) = False Then Text1.Text = "4000"  ' CLng(GetSetting("Window3D", "ButtonMap", "LThumbUp", "5000"))   ' left thumb dead zone
   If IsNumeric(Text2.Text) = False Then Text2.Text = "-4000"
   If IsNumeric(Text3.Text) = False Then Text3.Text = "-4000"
   If IsNumeric(Text4.Text) = False Then Text4.Text = "4000"
   If IsNumeric(Text5.Text) = False Then Text5.Text = "4000"
   If IsNumeric(Text6.Text) = False Then Text6.Text = "-4000"
   If IsNumeric(Text7.Text) = False Then Text7.Text = "-4000"
   If IsNumeric(Text8.Text) = False Then Text8.Text = "4000"
   isloaded = True
End Sub
Private Sub TextxChange(ByRef tb As TextBox)
   isloaded = False
   If IsNumeric(tb.Text) = False Then
      If tb.Tag = "" Then tb.Tag = 4000
      tb.Text = tb.Tag
   Else
      tb.Tag = tb.Text
   End If
   isloaded = True
End Sub
Private Sub Timer1_Timer()
   Dim x As Single
   Dim y As Single
   x = oldis.gamepad.sThumbLX
   y = oldis.gamepad.sThumbLY
   x = 65 + (44 * (x / 32767))
   y = 65 + (-44 * (y / 32767))
   DrawCircle Picture1, x, y, 20
   x = oldis.gamepad.sThumbRX
   y = oldis.gamepad.sThumbRY
   x = 65 + (44 * (x / 32767))
   y = 65 + (-44 * (y / 32767))
   DrawCircle Picture2, x, y, 20
End Sub
