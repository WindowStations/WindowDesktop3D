VERSION 5.00
Begin VB.Form frmLoad 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait..."
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
         Left            =   960
         TabIndex        =   1
         Top             =   1080
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   CreateRoundRectFromWindow Me
   CreateRoundRectFromWindow Frame1
   frmMain.SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, False, False
End Sub
Private Sub Form_Activate()
   WindowTransparency Me.hWnd, displayTransSettings, vbBlack
End Sub
