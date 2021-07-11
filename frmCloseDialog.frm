VERSION 5.00
Begin VB.Form frmCloseDialog 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3570
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   5355
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   720
         TabIndex        =   4
         Top             =   2040
         Width           =   1695
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Yes"
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
            Width           =   360
         End
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2640
         TabIndex        =   2
         Top             =   2040
         Width           =   1695
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No"
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
            Top             =   120
            Width           =   315
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Are you sure?"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   21.75
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
         Top             =   960
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmCloseDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   CreateRoundRectFromWindow Me
   CreateRoundRectFromWindow Frame1
   frmMain.SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, False, False
End Sub
Private Sub Frame2_Click()
   Unload frmMain
End Sub
Private Sub Frame8_Click()
   Unload Me
End Sub
Private Sub Label2_Click()
   Frame2_Click
End Sub
Private Sub Label8_Click()
   Frame8_Click
End Sub
Private Sub Frame8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Frame8.BackColor = &H808080
End Sub
Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Frame2.BackColor = &H808080
End Sub
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Frame8.BackColor <> &H404040 Then Frame8.BackColor = &H404040
   If Frame2.BackColor <> &H404040 Then Frame2.BackColor = &H404040
End Sub
