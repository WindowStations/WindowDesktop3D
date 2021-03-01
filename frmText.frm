VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmText 
   BorderStyle     =   0  'None
   Caption         =   "Gamepad settings"
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
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   9735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12735
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         Height          =   495
         Left            =   10440
         TabIndex        =   10
         Top             =   8760
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   9120
         TabIndex        =   9
         Top             =   8760
         Width           =   1215
      End
      Begin MSComctlLib.Slider traTextSpeed 
         Height          =   675
         Left            =   720
         TabIndex        =   1
         Top             =   2400
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   1191
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   100
         SelStart        =   40
         Value           =   40
      End
      Begin MSComctlLib.Slider traTextFade 
         Height          =   675
         Left            =   6600
         TabIndex        =   2
         Top             =   2400
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   1191
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   61
         SelStart        =   7
         Value           =   7
      End
      Begin MSComctlLib.Slider traTextSize 
         Height          =   675
         Left            =   720
         TabIndex        =   7
         Top             =   1080
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   1191
         _Version        =   393216
         LargeChange     =   1
         Min             =   8
         Max             =   72
         SelStart        =   8
         Value           =   8
      End
      Begin MSComctlLib.Slider traTextPosition 
         Height          =   675
         Left            =   6600
         TabIndex        =   8
         Top             =   1080
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   1191
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   4
         SelStart        =   4
         Value           =   4
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Pointer text display speed"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   720
         TabIndex        =   6
         Top             =   1920
         Width           =   5220
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Pointer text display position"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6600
         TabIndex        =   5
         Top             =   600
         Width           =   5220
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Pointer text display size"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   720
         TabIndex        =   4
         Top             =   600
         Width           =   5220
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Pointer text fade out time"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6600
         TabIndex        =   3
         Top             =   1920
         Width           =   5220
      End
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private isloaded  As Boolean
 
Private Sub Form_Load()
    traTextSize.Value = pointerTextSize
    traTextPosition.Value = pointerTextPosition
    traTextSpeed.Value = pointerTextSpeed
    traTextFade.Value = pointerTextFade
    Check1.Value = pointerMapDisplay
    isloaded = True
End Sub
Private Sub traTextSize_Scroll()
    If isloaded = False Then Exit Sub
    pointerTextSize = traTextSize.Value
End Sub
Private Sub traTextPosition_Scroll()
    If isloaded = False Then Exit Sub
    pointerTextPosition = traTextPosition.Value
End Sub

Private Sub traTextSpeed_Scroll()
    If isloaded = False Then Exit Sub
    pointerTextSpeed = traTextSpeed.Value
End Sub
Private Sub traTextFade_Click()
    If isloaded = False Then Exit Sub
    pointerTextFade = traWalkSpeed.Value
End Sub
Private Sub Check1_Click()
    If isloaded = False Then Exit Sub
    pointerMapDisplay = Check1.Value
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdApply_Click()
    pointerTextSize = traTextSize.Value
    pointerTextPosition = traTextPosition.Value
    pointerTextSpeed = traTextSpeed.Value
    pointerTextFade = traTextFade.Value
    pointerMapDisplay = Check1.Value
    SaveSetting "Window3D", "ButtonMap", "TextSize", CStr(pointerTextSize)
    SaveSetting "Window3D", "ButtonMap", "TextPosition", CStr(pointerTextPosition)
    SaveSetting "Window3D", "ButtonMap", "TextSpeed", CStr(pointerTextSpeed)
    SaveSetting "Window3D", "ButtonMap", "TextFade", CStr(pointerTextFade)
    SaveSetting "Window3D", "ButtonMap", "MapDisplay", CStr(pointerMapDisplay)
    Unload Me
End Sub
