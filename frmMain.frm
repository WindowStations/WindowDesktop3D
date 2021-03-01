VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   13020
   ClientLeft      =   -60
   ClientTop       =   -165
   ClientWidth     =   28800
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   868
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1920
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrload 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2040
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1320
      Top             =   1080
   End
   Begin VB.Timer tmrAutoRefreshIcons 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   1080
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      Height          =   780
      Left            =   120
      Picture         =   "frmMain.frx":1042
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   1
      Top             =   1080
      Width           =   510
   End
   Begin VB.PictureBox picMSG 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   960
      Left            =   0
      ScaleHeight     =   64
      ScaleMode       =   0  'User
      ScaleWidth      =   2000
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   30000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const GCL_STYLE                  As Long = (-26)
Private Const CS_DBLCLKS                 As Long = &H8
Private Const ERROR_DEVICE_NOT_CONNECTED As Long = 1167
Private Const ERROR_SUCCESS              As Long = 0
Private Const ERROR_EMPTY                As Long = 4306
Private Const EVENT_FOREGROUND           As Long = 3
Private Const GWL_EXSTYLE                As Long = -20
Private Const GWL_STYLE                  As Long = -16
Private Const GA_ROOT                    As Long = 2
Private Const HWND_TOPMOST               As Long = -1
Private Const HWND_DESKTOP               As Long = 0
Private Const HWND_TOP                   As Long = 0
Private Const HWND_BOTTOM                As Long = 1
Private Const HWND_NOTOPMOST             As Long = -2
Private Const KEYEVENTF_KEYDOWN          As Long = 0
Private Const KEYEVENTF_EXTENDEDKEY      As Long = &H1
Private Const KEYEVENTF_KEYUP            As Long = 2
Private Const MOUSEEVENTF_MOVE           As Long = 1
Private Const MOUSEEVENTF_LEFTDOWN       As Long = 2
Private Const MOUSEEVENTF_LEFTUP         As Long = 4
Private Const MOUSEEVENTF_RIGHTDOWN      As Long = 8
Private Const MOUSEEVENTF_RIGHTUP        As Long = 16
Private Const MOUSEEVENTF_MIDDLEDOWN     As Long = 32
Private Const MOUSEEVENTF_MIDDLEUP       As Long = 64
Private Const MOUSEEVENTF_XDOWN          As Long = 128
Private Const MOUSEEVENTF_XUP            As Long = 256
Private Const MOUSEEVENTF_WHEEL          As Long = 2048
Private Const MOUSEEVENTF_HWHEEL         As Long = 4096
Private Const MOUSEEVENTF_VIRTUALDESK    As Long = 16384
Private Const MOUSEEVENTF_ABSOLUTE       As Long = 32768
Private Const MOUSEEVENTF_WHEELROTATE    As Long = 120
Private Const SMTO_ABORTIFHUNG           As Long = &H2
Private Const SW_HIDE                    As Long = 0
Private Const SW_NORMAL                  As Long = 1
Private Const SW_SHOWMINIMIZED           As Long = 2
Private Const SW_SHOWMAXIMIZED           As Long = 3
Private Const SW_RESTORE                 As Long = 9
Private Const SM_FULLSCREEN              As Long = 65535
Private Const SWP_NOSIZE                 As Long = 1
Private Const SWP_NOMOVE                 As Long = 2
Private Const SWP_NOACTIVATE             As Long = 16
Private Const SWP_SHOWWINDOW             As Long = 64
Private Const SWP_NOOWNERZORDER          As Long = &H200
Private Const SWP_NOSENDCHANGING         As Long = &H400
Private Const WINEVENT_OUTOFCONTEXT      As Long = 0
Private Const WM_CLOSE                   As Long = 16
Private Const WS_EX_TOPMOST              As Long = 8
Private Const WS_EX_NOREDIRECTIONBITMAP  As Long = &H200000
Private Const WS_POPUPWINDOW             As Long = -2138570752
Private Const WS_EX_NOACTIVATE           As Long = &H8000000
Private Const NEGATIVE                   As Long = -1
Private Const HC_ACTION                  As Long = 0
Private Const HC_GETNEXT                 As Long = 1
Private Const WH_KEYBOARD_LL             As Long = 13
Private Const SW_SHOWNORMAL              As Long = 1
Private Const VK_SHIFT                   As Long = 16
Private Const VK_CONTROL                 As Long = 17
Private Const WM_KEYDOWN                 As Long = 256
Private Const WM_KEYUP                   As Long = 257
Private Const WM_SYSKEYDOWN              As Long = 260
Private Const WM_SYSKEYUP                As Long = 261
Private Const XINPUT_EXTRA_INFO          As Long = -32767
Private Const SM_CXVIRTUALSCREEN         As Long = 78
Private Const SM_CYVIRTUALSCREEN         As Long = 79
Private Const SM_CMONITORS               As Long = 80
Private Const SM_SAMEDISPLAYFORMAT       As Long = 81
Private Const WS_SIZEBOX                 As Long = &H40000
Private Const WS_THICKFRAME              As Long = &H40000
Private Const WS_MAXIMIZEBOX             As Long = &H10000
Private Const WS_MINIMIZEBOX             As Long = &H20000
Private Const QS_KEY                     As Long = 1
Private Const QS_MOUSEMOVE               As Long = 2
Private Const QS_MOUSEBUTTON             As Long = 4
Private Const QS_POSTMESSAGE             As Long = 8
Private Const QS_TIMER                   As Long = 16
Private Const QS_PAINT                   As Long = 32
Private Const QS_SENDMESSAGE             As Long = 64
Private Const QS_HOTKEY                  As Long = 128
Private Const QS_ALLPOSTMESSAGE          As Long = 256
Private Const QS_MOUSE                   As Long = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Private Const QS_INPUT                   As Long = (QS_MOUSE Or QS_KEY)
Private Const QS_ALLEVENTS               As Long = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)
Private Const QS_ALLINPUT                As Long = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
Private Const QS_ALL                     As Long = (QS_PAINT)
Private Const PM_NOREMOVE                As Long = &H0
Private Const PM_REMOVE                  As Long = &H1
Private Const WM_QUIT                    As Long = &H12
Private Const LWU_UNLOCK                 As Long = 0
Private Const IID_IImageList             As String = "{46EB5926-582E-4017-9FDF-E8998DAA0950}"
Private Const IID_IImageList2            As String = "{192B9D83-50FC-457B-90A0-2B82A8B5DAE1}"
Private Const SHIL_SMALL                 As Long = 1 ' 16x16
Private Const SHIL_LARGE                 As Long = 0 ' 32x32
Private Const SHIL_EXTRALARGE            As Long = 2 ' 48x48
Private Const SHIL_JUMBO                 As Long = 4 ' 256x256 (Vista+) & fails on XP or lower
Private Const E_INVALIDARG               As Long = &H80070057
Private Const DT_WORDBREAK               As Long = &H10
Private Const VER_PLATFORM_WIN32_WINDOWS    As Long = 1
Private Const TH32CS_SNAPPROCESS            As Long = &H2
Private Const PROCESS_VM_READ               As Long = 16
Private Const PROCESS_QUERY_INFORMATION     As Long = 1024
Private Const PROCESS_TERMINATE             As Long = &H1
Private Const MAX_PATH                      As Long = 260
Private Const CSIDL_SHELLNEW                As Long = 21
Private Const CSIDL_DESKTOP                 As Long = &H0
Private Const CSIDL_INTERNET                As Long = &H1
Private Const CSIDL_PROGRAMS                As Long = &H2
Private Const CSIDL_CONTROLS                As Long = &H3
Private Const CSIDL_PRINTERS                As Long = &H4
Private Const CSIDL_PERSONAL                As Long = &H5
Private Const CSIDL_FAVORITES               As Long = &H6
Private Const CSIDL_STARTUP                 As Long = &H7
Private Const CSIDL_RECENT                  As Long = &H8
Private Const CSIDL_SENDTO                  As Long = &H9
Private Const CSIDL_BITBUCKET               As Long = &HA
Private Const CSIDL_STARTMENU               As Long = &HB
Private Const CSIDL_MYDOCUMENTS             As Long = &HC
Private Const CSIDL_MYMUSIC                 As Long = &HD
Private Const CSIDL_MYVIDEO                 As Long = &HE
Private Const CSIDL_DESKTOPDIRECTORY        As Long = &H10
Private Const CSIDL_DRIVES                  As Long = &H11
Private Const CSIDL_NETWORK                 As Long = &H12
Private Const CSIDL_NETHOOD                 As Long = &H13
Private Const CSIDL_FONTS                   As Long = &H14
Private Const CSIDL_TEMPLATES               As Long = &H15
Private Const CSIDL_COMMON_STARTMENU        As Long = &H16
Private Const CSIDL_COMMON_PROGRAMS         As Long = &H17
Private Const CSIDL_COMMON_STARTUP          As Long = &H18
Private Const CSIDL_COMMON_DESKTOPDIRECTORY As Long = &H19
Private Const CSIDL_APPDATA                 As Long = &H1A
Private Const CSIDL_PRINTHOOD               As Long = &H1B
Private Const CSIDL_LOCAL_APPDATA           As Long = &H1C
Private Const CSIDL_ALTSTARTUP              As Long = &H1D
Private Const CSIDL_COMMON_ALTSTARTUP       As Long = &H1E
Private Const CSIDL_COMMON_FAVORITES        As Long = &H1F
Private Const CSIDL_INTERNET_CACHE          As Long = &H20
Private Const CSIDL_COOKIES                 As Long = &H21
Private Const CSIDL_HISTORY                 As Long = &H22
Private Const CSIDL_COMMON_APPDATA          As Long = &H23
Private Const CSIDL_WINDOWS                 As Long = &H24
Private Const CSIDL_SYSTEM                  As Long = &H25
Private Const CSIDL_PROGRAM_FILES           As Long = &H26
Private Const CSIDL_MYPICTURES              As Long = &H27
Private Const CSIDL_PROFILE                 As Long = &H28
Private Const CSIDL_SYSTEMX86               As Long = &H29
Private Const CSIDL_PROGRAM_FILESX86        As Long = &H2A
Private Const CSIDL_PROGRAM_FILES_COMMON    As Long = &H2B
Private Const CSIDL_PROGRAM_FILES_COMMONX86 As Long = &H2C
Private Const CSIDL_COMMON_TEMPLATES        As Long = &H2D
Private Const CSIDL_COMMON_DOCUMENTS        As Long = &H2E
Private Const CSIDL_COMMON_ADMINTOOLS       As Long = &H2F
Private Const CSIDL_ADMINTOOLS              As Long = &H30
Private Const CSIDL_CONNECTIONS             As Long = &H31
Private Const CSIDL_COMMON_MUSIC            As Long = &H35
Private Const CSIDL_COMMON_PICTURES         As Long = &H36
Private Const CSIDL_COMMON_VIDEO            As Long = &H37
Private Const CSIDL_RESOURCES               As Long = &H38
Private Const CSIDL_RESOURCES_LOCALIZED     As Long = &H39
Private Const CSIDL_COMMON_OEM_LINKS        As Long = &H3A
Private Const CSIDL_CDBURN_AREA             As Long = &H3B
Private Const CSIDL_COMPUTERSNEARME         As Long = &H3D
Private Const CSIDL_FLAG_PER_USER_INIT      As Long = &H800
Private Const CSIDL_FLAG_NO_ALIAS           As Long = &H1000
Private Const CSIDL_FLAG_DONT_VERIFY        As Long = &H4000
Private Const CSIDL_FLAG_CREATE             As Long = &H8000
Private Const CSIDL_FLAG_MASK               As Long = &HFF00
Const RDW_INVALIDATE As Long = 1 'Invalidates the redraw area.
Const RDW_INTERNALPAINT As Long = 2 'A WM_PAINT message is posted to the window even if it is not invalid.
Const RDW_ERASE As Long = 4 '  The background of the redraw area is erased before drawing. RDW_INVALIDATE must also be specified.
Const RDW_VALIDATE As Long = 8 'Validates the redraw area.
Const RDW_NOINTERNALPAINT As Long = 16 'Prevents any pending WM_PAINT messages that were generated internally or by this function. WM_PAINT messages will still be generated for invalid areas.
Const RDW_NOERASE As Long = 32 'Prevents the background of the redraw area from being erased.
Const RDW_NOCHILDREN As Long = 64 'Redraw operation excludes child windows if present in the redraw area.
Const RDW_ALLCHILDREN As Long = 128 'Redraw operation includes child windows if present in the redraw area.
Const RDW_UPDATENOW As Long = 256 'Updates the specified redraw area immediately.
Const RDW_ERASENOW As Long = 512 'Erases the specified redraw area immediately.
Const RDW_FRAME As Long = 1024 'Updates the nonclient area if included in the redraw area. RDW_INVALIDATE must also be specified.
Const RDW_NOFRAME As Long = 2048 'Prevents the nonclient area from being redrawn if it is part of the redraw area. RDW_VALIDATE must also be specified.
'    Const HWND_DESKTOP As Long = 0
'    Private Type RECT
'         rLeft As Long
'         rTop As Long
'         rRight As Long
'         rBottom As Long
'    End Type
'Private Const SMTO_ABORTIFHUNG                    As Long = &H2
Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type
Private Type WINNAME
    lpText As String
    lpClass As String
End Type
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type QUE_INPUT
    Adown As Boolean
    Aup As Boolean
    bdown As Boolean
    bup As Boolean
    xdown As Boolean
    xup As Boolean
    ydown As Boolean
    yup As Boolean
    dleftdown As Boolean
    dleftup As Boolean
    drightdown As Boolean
    drightup As Boolean
    dupdown As Boolean
    dupup As Boolean
    ddowndown As Boolean
    ddownup As Boolean
    lbumperdown As Boolean
    lbumperup As Boolean
    rbumperdown As Boolean
    rbumperup As Boolean
    lstickdown As Boolean
    lstickup As Boolean
    rstickdown As Boolean
    rstickup As Boolean
    backdown As Boolean
    backup As Boolean
    startdown As Boolean
    startup As Boolean
End Type
Private Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    Point As POINTAPI
End Type
Private Enum Keys
    VK_None = 0
    VK_LButton = 1
    VK_RButton = 2
    VK_Cancel = 3
    VK_MButton = 4
    VK_XButton1 = 5
    VK_XButton2 = 6
    VK_LButton_XButton2 = 7
    vk_back = 8
    vk_Tab = 9
    VK_LineFeed = 10
    VK_LButton_LineFeed = 11
    VK_Clear = 12
    vk_return = 13
    VK_RButton_Clear = 14
    VK_RButton_Return = 15
    vk_ShiftKey = 16
    VK_controlkey = 17
    VK_MENU = 18
    VK_Pause = 19
    VK_CAPITAL = 20
    VK_KanaMode = 21
    VK_RButton_Capital = 22
    VK_JunjaMode = 23
    VK_FinalMode = 24
    VK_HanjaMode = 25
    VK_RButton_FinalMode = 26
    vk_Escape = 27
    VK_IMEConvert = 28
    VK_IMENonconvert = 29
    VK_IMEAceept = 30
    VK_IMEModeChange = 31
    VK_Space = 32
    VK_PageUp = 33
    VK_Next = 34
    VK_End = 35
    VK_Home = 36
    vk_Left = 37
    vk_up = 38
    vk_Right = 39
    vk_down = 40
    VK_Select = 41
    VK_Print = 42
    VK_Execute = 43
    VK_PrintScreen = 44
    VK_Insert = 45
    VK_delete = 46
    VK_Help = 47
    vk_d0 = 48
    vk_d1 = 49
    vk_d2 = 50
    vk_d3 = 51
    vk_d4 = 52
    vk_d5 = 53
    vk_d6 = 54
    vk_d7 = 55
    vk_d8 = 56
    vk_d9 = 57
    VK_RButton_D8 = 58
    VK_RButton_D9 = 59
    VK_MButton_D8 = 60
    VK_MButton_D9 = 61
    VK_XButton2_D8 = 62
    VK_XButton2_D9 = 63
    VK_64 = 64
    vk_a = 65
    vk_b = 66
    vk_c = 67
    vk_d = 68
    vk_e = 69
    vk_f = 70
    vk_g = 71
    vk_h = 72
    vk_i = 73
    vk_j = 74
    vk_k = 75
    vk_l = 76
    vk_m = 77
    vk_n = 78
    vk_o = 79
    vk_p = 80
    vk_q = 81
    vk_r = 82
    vk_s = 83
    vk_t = 84
    vk_u = 85
    vk_v = 86
    vk_w = 87
    vk_x = 88
    vk_y = 89
    vk_z = 90
    VK_LWIN = 91
    VK_RWIN = 92
    vk_Apps = 93
    VK_RButton_RWin = 94
    VK_Sleep = 95
    VK_NumPad0 = 96
    VK_NumPad1 = 97
    VK_NumPad2 = 98
    VK_NumPad3 = 99
    VK_NumPad4 = 100
    VK_NumPad5 = 101
    VK_NumPad6 = 102
    VK_NumPad7 = 103
    VK_NumPad8 = 104
    VK_NumPad9 = 105
    VK_Multiply = 106
    VK_Add = 107
    VK_Separator = 108
    VK_Subtract = 109
    VK_Decimal = 110
    VK_Divide = 111
    VK_F1 = 112
    VK_F2 = 113
    VK_F3 = 114
    VK_F4 = 115
    VK_F5 = 116
    VK_F6 = 117
    VK_F7 = 118
    VK_F8 = 119
    VK_F9 = 120
    VK_F10 = 121
    VK_F11 = 122
    VK_F12 = 123
    VK_F13 = 124
    VK_F14 = 125
    VK_F15 = 126
    VK_F16 = 127
    VK_F17 = 128
    VK_F18 = 129
    VK_F19 = 130
    VK_F20 = 131
    VK_F21 = 132
    VK_F22 = 133
    VK_F23 = 134
    VK_F24 = 135
    VK_Back_F17 = 136
    VK_Back_F18 = 137
    VK_Back_F19 = 138
    VK_Back_F20 = 139
    VK_Back_F21 = 140
    VK_Back_F22 = 141
    VK_Back_F23 = 142
    VK_Back_F24 = 143
    VK_NumLock = 144
    VK_Scroll = 145
    VK_RButton_NumLock = 146
    VK_RButton_Scroll = 147
    VK_MButton_NumLock = 148
    VK_MButton_Scroll = 149
    VK_XButton2_NumLock = 150
    VK_XButton2_Scroll = 151
    VK_Back_NumLock = 152
    VK_Back_Scroll = 153
    VK_LineFeed_NumLock = 154
    VK_LineFeed_Scroll = 155
    VK_Clear_NumLock = 156
    VK_Clear_Scroll = 157
    VK_RButton_Clear_NumLock = 158
    VK_RButton_Clear_Scroll = 159
    VK_LShiftKey = 160
    VK_RShiftKey = 161
    VK_LControlKey = 162
    VK_RControlKey = 163
    VK_LMenu = 164
    VK_RMenu = 165
    VK_BrowserBack = 166
    VK_BrowserForward = 167
    VK_BrowserRefresh = 168
    VK_BrowserStop = 169
    VK_BrowserSearch = 170
    VK_BrowserFavorites = 171
    VK_BrowserHome = 172
    VK_VolumeMute = 173
    VK_VolumeDown = 174
    VK_VolumeUp = 175
    VK_MediaNextTrack = 176
    VK_MediaPreviousTrack = 177
    VK_MediaStop = 178
    VK_MediaPlayPause = 179
    VK_LaunchMail = 180
    VK_SelectMedia = 181
    VK_LaunchApplication1 = 182
    VK_LaunchApplication2 = 183
    VK_Back_MediaNextTrack = 184
    VK_Back_MediaPreviousTrack = 185
    VK_OemSemiColon = 186
    vk_oemplus = 187
    vk_oemcomma = 188
    VK_OemMinus = 189
    vk_oemperiod = 190
    vk_oemquestion = 191
    VK_Oemtilde = 192
    VK_LButton_Oemtilde = 193
    VK_RButton_Oemtilde = 194
    VK_Cancel_Oemtilde = 195
    VK_MButton_Oemtilde = 196
    VK_XButton1_Oemtilde = 197
    VK_XButton2_Oemtilde = 198
    VK_LButton_XButton2_Oemtilde = 199
    VK_Back_Oemtilde = 200
    VK_Tab_Oemtilde = 201
    VK_LineFeed_Oemtilde = 202
    VK_LButton_LineFeed_Oemtilde = 203
    VK_Clear_Oemtilde = 204
    VK_Return_Oemtilde = 205
    VK_RButton_Clear_Oemtilde = 206
    VK_RButton_Return_Oemtilde = 207
    VK_ShiftKey_Oemtilde = 208
    VK_ControlKey_Oemtilde = 209
    VK_Menu_Oemtilde = 210
    VK_Pause_Oemtilde = 211
    VK_Capital_Oemtilde = 212
    VK_KanaMode_Oemtilde = 213
    VK_RButton_Capital_Oemtilde = 214
    VK_JunjaMode_Oemtilde = 215
    VK_FinalMode_Oemtilde = 216
    VK_HanjaMode_Oemtilde = 217
    VK_RButton_FinalMode_Oemtilde = 218
    VK_OemOpenBrackets = 219
    VK_Oem5 = 220
    VK_OemCloseBracket = 221
    VK_oemApostrophe = 222
    VK_Oem8 = 223
    VK_Space_Oemtilde = 224
    VK_PageUp_Oemtilde = 225
    VK_OemBackslash = 226
    VK_LButton_OemBackslash = 227
    VK_Home_Oemtilde = 228
    VK_ProcessKey = 229
    VK_MButton_OemBackslash = 230
    VK_Packet = 231
    VK_Down_Oemtilde = 232
    VK_Select_Oemtilde = 233
    VK_Back_OemBackslash = 234
    VK_Tab_OemBackslash = 235
    VK_PrintScreen_Oemtilde = 236
    VK_Back_ProcessKey = 237
    VK_Clear_OemBackslash = 238
    VK_Back_Packet = 239
    VK_D0_Oemtilde = 240
    VK_D1_Oemtilde = 241
    VK_ShiftKey_OemBackslash = 242
    VK_ControlKey_OemBackslash = 243
    VK_D4_Oemtilde = 244
    VK_ShiftKey_ProcessKey = 245
    VK_Attn = 246
    VK_Crsel = 247
    VK_Exsel = 248
    VK_EraseEof = 249
    VK_Play = 250
    VK_Zoom = 251
    VK_NoName = 252
    VK_Pa1 = 253
    VK_OemClear = 254
    VK_LButton_OemClear = 255
End Enum
Private Enum HWND_
    TOPMOST = HWND_TOPMOST
    bottom = HWND_BOTTOM
    top = HWND_TOP
    NOTOPMOST = HWND_NOTOPMOST
End Enum
Private Enum MOUSEEVENTF_
    LeftDown = MOUSEEVENTF_LEFTDOWN
    LeftUp = MOUSEEVENTF_LEFTUP
    LeftClick = MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP
    LeftDoubleClick = MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP + MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP
    MiddleDown = MOUSEEVENTF_MIDDLEDOWN
    middleUp = MOUSEEVENTF_MIDDLEUP
    MiddleClick = MOUSEEVENTF_MIDDLEDOWN + MOUSEEVENTF_MIDDLEUP
    MiddleDoubleClick = MOUSEEVENTF_MIDDLEDOWN + MOUSEEVENTF_MIDDLEUP + MOUSEEVENTF_MIDDLEDOWN + MOUSEEVENTF_MIDDLEUP
    Move = MOUSEEVENTF_MOVE
    MoveAbsolute = MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_MOVE
    RightDown = MOUSEEVENTF_RIGHTDOWN
    RightUp = MOUSEEVENTF_RIGHTUP
    RightClick = MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP
    RightDoubleClick = MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP + MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP
    VirtualDesk = MOUSEEVENTF_VIRTUALDESK
    WHEEL = MOUSEEVENTF_WHEEL
    xdown = MOUSEEVENTF_XDOWN
    xup = MOUSEEVENTF_XUP
    xclick = MOUSEEVENTF_XDOWN + MOUSEEVENTF_XUP
    xDoubleClick = MOUSEEVENTF_XDOWN + MOUSEEVENTF_XUP + MOUSEEVENTF_XDOWN + MOUSEEVENTF_XUP
End Enum
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128
End Type
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Private Type PICTDESC
    cbSize As Long
    PicType As Long
    hImage As Long
    Data1 As Long
    Data2 As Long
End Type
Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type
Private Declare Function apiRedrawWindow Lib "user32" Alias "RedrawWindow" (ByVal hWnd As Long, ByVal lprcUpdate As Boolean, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function apiIsIconic Lib "user32" Alias "IsIconic" (ByVal hWnd As Long) As Long
Private Declare Function apiProcess32First Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, ByRef lppe As PROCESSENTRY32) As Long
Private Declare Function apiProcess32Next Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, ByRef lppe As PROCESSENTRY32) As Long
Private Declare Function apiCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal Handle As Long) As Long
Private Declare Function apiOpenProcess Lib "kernel32" Alias "OpenProcess" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function apiCreateToolhelp32Snapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function apiGetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function apiTerminateProcess Lib "kernel32" Alias "TerminateProcess" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function apiGetWindowThreadProcessId Lib "user32" Alias "GetWindowThreadProcessId" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function apiGetModuleFileNameExA Lib "psapi" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function apiEnumProcessModules Lib "psapi" Alias "EnumProcessModules" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function apiEnumProcesses Lib "psapi" Alias "EnumProcesses" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function apiMoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Byte, ByVal Source As Long, ByVal Length As Long) As Long
Private Declare Function apilstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function apiGetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function apiGetFileVersionInfo Lib "Version" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, ByRef lpData As Byte) As Long
Private Declare Function apiGetFileVersionInfoSize Lib "Version" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, ByRef lpdwHandle As Long) As Long
Private Declare Function apiVerQueryValueByteLong Lib "Version" Alias "VerQueryValueA" (ByRef pBlock As Byte, ByVal lpSubBlock As String, ByRef lplpBuffer As Long, ByRef puLen As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (ByRef pPictDesc As PICTDESC, ByRef RefIID As GUID, ByVal fPictureOwnsHandle As Long, ByRef ppvObj As StdPicture) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function SHGetImageListXP Lib "shell32.dll" Alias "#727" (ByVal iImageList As Long, ByRef rIID As Long, ByRef ppv As Any) As Long
Private Declare Function SHGetImageList Lib "shell32.dll" (ByVal iImageList As Long, ByRef rIID As Long, ByRef ppv As Any) As Long
Private Declare Function IIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByRef lpiid As Any) As Long
Private Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal hIML As Long, ByVal i As Long, ByVal flags As Long) As Long
Private Declare Function apiSetForegroundWindow Lib "user32" Alias "SetForegroundWindow" (ByVal hWnd As Long) As Long
Private Declare Function DrawIconEx Lib "USER32.DLL" (ByVal hdc As Long, ByVal XLeft As Long, ByVal YTop As Long, ByVal hIcon As Long, ByVal CXWidth As Long, ByVal CYWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "USER32.DLL" (ByVal hIcon As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function apiGetForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As Long
Private Declare Function apiCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Byte, ByRef lpSource As XINPUT_STATE, ByVal cbCopy As Long) As Long
Private Declare Function apiCopyMemoryType Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As XINPUT_STATE, ByRef lpSource As XINPUT_STATE, ByVal cbCopy As Long) As Long
Private Declare Function apiGetTickCount Lib "kernel32" Alias "GetTickCount" () As Long
Private Declare Function apiExitProcess Lib "kernel32" Alias "ExitProcess" (ByVal uExitCode As Long) As Long
Private Declare Function apiQueryPerformanceCounter Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef lpPerformanceCount As Currency) As Long
Private Declare Function apiQueryPerformanceFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef lpFrequency As Currency) As Long
Private Declare Function apiPeekMessage Lib "user32" Alias "PeekMessageA" (ByRef lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function apiGetMessage Lib "user32" Alias "GetMessageA" (ByRef lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function apiTranslateMessage Lib "user32" Alias "TranslateMessage" (ByRef lpMsg As Msg) As Long
Private Declare Function apiDispatchMessage Lib "user32" Alias "DispatchMessageA" (ByRef lpMsg As Msg) As Long
Private Declare Function apiGetQueueStatus Lib "user32" Alias "GetQueueStatus" (ByVal fuFlags As Long) As Long
Private Declare Function apiSetCursorPos Lib "user32" Alias "SetCursorPos" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function apiGetCursorPos Lib "user32" Alias "GetCursorPos" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function apiSetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function apiGetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function apiGetMessageExtraInfo Lib "user32" Alias "GetMessageExtraInfo" () As Long
Private Declare Function apiGetSystemMetrics Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Private Declare Function apikeybd_event Lib "user32" Alias "keybd_event" (ByVal vKey As Long, ByVal bScan As Long, ByVal dwFlags As Long, ByVal dwExtraInfo As Long) As Long
Private Declare Function apimouse_event Lib "user32" Alias "mouse_event" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long) As Boolean
Private Declare Function apiOemKeyScan Lib "user32" Alias "OemKeyScan" (ByVal wOemChar As Long) As Long
Private Declare Function apiSetTimer1 Lib "user32" Alias "SetTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function apiKillTimer1 Lib "user32" Alias "KillTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function apiSetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function apiShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function apiMoveWindow Lib "user32" Alias "MoveWindow" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Boolean) As Boolean
Private Declare Function apiIsWindow Lib "user32" Alias "IsWindow" (ByVal hWnd As Long) As Long
Private Declare Function apiGetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As Long
Private Declare Function apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function apiGetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function apiSetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function apiWindowFromPoint Lib "user32" Alias "WindowFromPoint" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function apiGetAncestor Lib "user32" Alias "GetAncestor" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Private Declare Function apiSendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, ByRef lpdwResult As Long) As Long
Private Declare Function apiGetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function apiGetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function apiGetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function apiIsZoomed Lib "user32" Alias "IsZoomed" (ByVal hWnd As Long) As Long
Private Declare Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function apiLockWindowUpdate Lib "user32" Alias "LockWindowUpdate" (ByVal hwndLock As Long) As Long
Private Declare Function apiIsWindowVisible Lib "user32" Alias "IsWindowVisible" (ByVal hWnd As Long) As Long
'Private Declare Function apiSetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
'Private Declare Function apiSetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function apiGetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function apiShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Private Declare Function apiGetSystemMetrics Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Public WithEvents TMRPoll          As Timer
Attribute TMRPoll.VB_VarHelpID = -1
Private WithEvents TMRPollBatteries As Timer
Attribute TMRPollBatteries.VB_VarHelpID = -1
Public WithEvents xinputClass      As clsXInput
Attribute xinputClass.VB_VarHelpID = -1
Public WithEvents itemclass         As clsItems
Attribute itemclass.VB_VarHelpID = -1
Public WithEvents buttonClass       As clsButton
Attribute buttonClass.VB_VarHelpID = -1
Public WithEvents ProjectileClass   As clsProjectile
Attribute ProjectileClass.VB_VarHelpID = -1
Public WithEvents mapClass          As clsMap
Attribute mapClass.VB_VarHelpID = -1
Public consoleClass                 As clsConsole
Public spriteClass                  As clsSprite
Public shadowClass                  As clsShadow
Public soundClass                   As clsSoundEffects
Private eventtime                   As Long
Private leftmousedown               As Boolean ': leftmousedown = False
Private shiftkeydown                As Long ': shiftkeydown = False
Private isloaded                    As Boolean ': isloaded = False
Private ispainted                   As Boolean ': ispainted = False
Private xenable                    As Boolean
Private rollingscroll               As Long ': rollingscroll = 0
Private IsXINPUTSupported           As Boolean ': IsXINPUTSupported = False
Private cwnd                        As Long ': cwnd = 0
Private qi                          As QUE_INPUT
Private active_input_user           As Long
Private lasttick                    As Long
Private hDCScreen                   As Long
Private rectCurrent                 As RECT
Private xinpforward                 As Boolean
Private xinpbackward                As Boolean
Private xinpleft                    As Boolean
Private xinpright                   As Boolean
Private bCancel                     As Boolean
Private SinD                        As Single
Private CosD                        As Single
Private SinA                        As Single
Private CosA                        As Single
Private RetKey                      As Long
Private QF                          As Currency
Private OldQC                       As Currency
Private MinCur                      As Currency
Private MaxCur                      As Currency
Private QTimeVal                    As Double
Private MouseSensX                  As Single
Private MouseSensY                  As Single
Private ScrCenterX                  As Long
Private ScrCenterY                  As Long
Public isconnected                  As Boolean
Private D3D                         As Direct3D9
Private d3dpp                       As D3DPRESENT_PARAMETERS
Private lmousedown                  As Boolean
Private oldpoint                    As POINTAPI
Private pullwnd                     As Long
Private foo                         As Boolean
Private foohittest                  As Boolean
Private sHeight                     As Long
Private lastselect                  As Long
Private k                           As Long
Private oldfls As String
Private minwinds() As Long
Private lastmodifierstamp As Long
Private pointerAcc As Boolean
Private moveleft As Double
Private moveright As Double
Private moveforward As Double
Private movebackward As Double
Private showdesktopicons As Boolean



Private Sub Form_Initialize()
    '   frmLoad.show
    SetWindowPos Me.hWnd, HWND_.TOPMOST, 0, 0, 0, 0, True, True
    WindowTransparency Me.hWnd, displayTrans3D, vbBlack
End Sub
Friend Sub reload()
On Error Resume Next
    D3DInit Me.hWnd
    QTimeReset 0
    buttonClass.Initialize
    mapClass.Load MapFileName
    itemclass.Initialize
    shadowClass.Initialize
    spriteClass.Initialize
    PhysInit
    consoleClass.Initialize
    soundClass.Initialize (Me.hWnd)
End Sub
Private Sub Form_Load()
    On Error Resume Next
    resizetoPrimaryscreen
    keymapA = GetSetting("Window3D", "ButtonMap", "AButton", "")   ' "Action button (Left Click)"
    keymapMenu = GetSetting("Window3D", "ButtonMap", "Menu", "")   ' "Context Menu button (Right Click)"
    keymapB = GetSetting("Window3D", "ButtonMap", "BButton", "")   ' "Cancel Button(Escape)"
    keymapY = GetSetting("Window3D", "ButtonMap", "YButton", "")   ' "Toggle 2D/3D view"
    keymapX = GetSetting("Window3D", "ButtonMap", "XButton", "")   ' "Close/Terminate Window button"
    keymapLeftBumper = GetSetting("Window3D", "ButtonMap", "LeftBumper", "")  '  "Minimize Window button"
    keymapRightBumper = GetSetting("Window3D", "ButtonMap", "RightBumper", "")  '  "Maximize/Restore Window toggle"
    keymapLeftStick = GetSetting("Window3D", "ButtonMap", "LeftStick", "")   ' "Push Window to bottom"
    keymapRightStick = GetSetting("Window3D", "ButtonMap", "RightStick", "")   ' "Pull Window to top"
    keymapDLeft = GetSetting("Window3D", "ButtonMap", "DLeft", "")   ' "Left Arrow key"
    keymapDRight = GetSetting("Window3D", "ButtonMap", "DRight", "")   ' "Right Arrow key"
    keymapDUp = GetSetting("Window3D", "ButtonMap", "DUp", "")   ' "Up Arrow key"
    keymapDDown = GetSetting("Window3D", "ButtonMap", "DDown", "")   ' "Down Arrow key"
    keymapChange = GetSetting("Window3D", "ButtonMap", "Change", "")   ' "Modifier key (alternate function/Settings)"
    keymapLThumbUp = CLng(GetSetting("Window3D", "ButtonMap", "LThumbUp", "4000"))   ' left thumb dead zone
    keymapLThumbDown = CLng(GetSetting("Window3D", "ButtonMap", "LThumbDown", "-4000"))
    keymapLThumbLeft = CLng(GetSetting("Window3D", "ButtonMap", "LThumbLeft", "-4000"))
    keymapLThumbRight = CLng(GetSetting("Window3D", "ButtonMap", "LThumbRight", "4000"))
    keymapRThumbUp = CLng(GetSetting("Window3D", "ButtonMap", "RThumbUp", "4000"))   ' right thumb dead zone
    keymapRThumbDown = CLng(GetSetting("Window3D", "ButtonMap", "RThumbDown", "-4000"))
    keymapRThumbLeft = CLng(GetSetting("Window3D", "ButtonMap", "RThumbLeft", "-4000"))
    keymapRThumbRight = CLng(GetSetting("Window3D", "ButtonMap", "RThumbRight", "4000"))
    keymapDisablegamepad = CLng(GetSetting("Window3D", "ButtonMap", "DisableGamepad", "0"))
    If keymapA = "" Then keymapA = "1": SaveSetting "Window3D", "ButtonMap", "AButton", 1
    If keymapMenu = "" Then keymapMenu = "2": SaveSetting "Window3D", "ButtonMap", "Menu", 2
    If keymapB = "" Then keymapB = "3": SaveSetting "Window3D", "ButtonMap", "BButton", 3
    If keymapY = "" Then keymapY = "4": SaveSetting "Window3D", "ButtonMap", "YButton", 4
    If keymapX = "" Then keymapX = "5": SaveSetting "Window3D", "ButtonMap", "XButton", 5
    If keymapLeftBumper = "" Then keymapLeftBumper = "6": SaveSetting "Window3D", "ButtonMap", "LeftBumper", 6
    If keymapRightBumper = "" Then keymapRightBumper = "7": SaveSetting "Window3D", "ButtonMap", "RightBumper", 7
    If keymapLeftStick = "" Then keymapLeftStick = "8": SaveSetting "Window3D", "ButtonMap", "LeftStick", 8
    If keymapRightStick = "" Then keymapRightStick = "9": SaveSetting "Window3D", "ButtonMap", "RightStick", 9
    If keymapDLeft = "" Then keymapDLeft = "10": SaveSetting "Window3D", "ButtonMap", "DLeft", 10
    If keymapDRight = "" Then keymapDRight = "11": SaveSetting "Window3D", "ButtonMap", "DRight", 11
    If keymapDUp = "" Then keymapDUp = "12": SaveSetting "Window3D", "ButtonMap", "DUp", 12
    If keymapDDown = "" Then keymapDDown = "13": SaveSetting "Window3D", "ButtonMap", "DDown", 13
    If keymapChange = "" Then keymapChange = "14": SaveSetting "Window3D", "ButtonMap", "Change", 14
    pointerMaxPointerSpeed = CLng(GetSetting("Window3D", "Pointer", "MaxPointerSpeed", "6"))
    pointerMaxPointerAcceleration = CLng(GetSetting("Window3D", "Pointer", "MaxPointerAcceleration", "20"))
    pointerMaxWheelSpeed = CLng(GetSetting("Window3D", "Pointer", "MaxWheelSpeed", "10"))
    pointerMaxWheelAcceleration = CLng(GetSetting("Window3D", "Pointer", "MaxWheelAcceleration", "20"))
    pointerMaxPOVSpeed = CLng(GetSetting("Window3D", "Pointer", "MaxPOVSpeed", "10"))
    pointerMaxPOVAcceleration = CLng(GetSetting("Window3D", "Pointer", "MaxPOVAcceleration", "40"))
    pointerMaxWalkSpeed = CLng(GetSetting("Window3D", "Pointer", "MaxWalkSpeed", "1"))
    pointerMaxWalkAcceleration = CLng(GetSetting("Window3D", "Pointer", "MaxWalkAcceleration", "4"))
    pointerDisable2D = CLng(GetSetting("Window3D", "Pointer", "Disable2D", "0"))
    pointerDisable3D = CLng(GetSetting("Window3D", "Pointer", "Disable3D", "0"))
    displayDelay = CLng(GetSetting("Window3D", "Display", "Delay", "500"))
    displaySpeed = CLng(GetSetting("Window3D", "Display", "Speed", "1"))
    displayFade = CLng(GetSetting("Window3D", "Display", "Fade", "3000"))
    displayPosition = CLng(GetSetting("Window3D", "Display", "Position", "1000"))
    displayTrans3D = CLng(GetSetting("Window3D", "Display", "3D", "255"))
    displayTransSettings = CInt(GetSetting("Window3D", "Display", "Settings", "250"))
    displayHide = CLng(GetSetting("Window3D", "Display", "Hide", "0"))
    directxTexFilters = CLng(GetSetting("Window3D", "DirectX", "TexFilters", "2"))
    directxVSync = CBool(GetSetting("Window3D", "DirectX", "VSync", "0"))
    directxQuant = 1 / CDbl(GetSetting("Window3D", "DirectX", "Quant", "200"))
    directxGravity = 1 / CDbl(GetSetting("Window3D", "DirectX", "Gravity", "-1000")) '1 / 180
    directxTexFIndex = CLng(GetSetting("Window3D", "DirectX", "TexFIndex", "4")) '4
    directxAnisotropy = CLng(GetSetting("Window3D", "DirectX", "Anisotropy", "8")) '2 ^ (directxTexFIndex - 1)
    Dim pi60 As Single
    pi60 = 60 * Pi / 180
    directxFovY = CSng(GetSetting("Window3D", "DirectX", "FovY", CStr(pi60)))
    Dim asra As Single
    asra = (Screen.Width / Screen.TwipsPerPixelX) / (Screen.Height / Screen.TwipsPerPixelY)
    directxAspect = CSng(GetSetting("Window3D", "DirectX", "Aspect", CStr(asra)))
    soundxMute = CLng(GetSetting("Window3D", "Soundeffects", "Mute", "0"))
    'initialize classes for 3d directx
    MapFileName = App.Path & "\3DEngine\Maps\" & "Quits" & ".map"
    MouseSensX = 0.00003 * 20
    MouseSensY = 0.00003 * 20 * ((0 <> 0) * 2 + 1)
    Set itemclass = New clsItems
    Set buttonClass = New clsButton
    Set ProjectileClass = New clsProjectile
    Set mapClass = New clsMap
    Set consoleClass = New clsConsole
    Set spriteClass = New clsSprite
    Set shadowClass = New clsShadow
    Set TMRPoll = New Timer
    Set TMRPollBatteries = New Timer
    Set xinputClass = New clsXInput
    Set soundClass = New clsSoundEffects
    CenterCursor
    D3DInit Me.hWnd
    QTimeReset 0
    buttonClass.Initialize
    mapClass.Load MapFileName
    itemclass.Initialize
    shadowClass.Initialize
    spriteClass.Initialize
    PhysInit
    consoleClass.Initialize
    soundClass.Initialize (Me.hWnd)
    'set up performance counter
    OldTime = QTime
    NowTime = OldTime
    SaveGame App.Path & "\Game0.sav"
    SaveGame App.Path & "\Game1.sav"
    WheelHook Me.hWnd
    TMRPoll.Interval = 1
    TMRPollBatteries.Interval = 1
    TMRPoll.Enabled = True
    TMRPollBatteries.Enabled = True
    If keymapDisablegamepad = 1 Then
        xinputClass.Disable
    Else
        xinputClass.Enable
    End If
    If pointerDisable3D = 1 And Me.Visible = True Then Me.Visible = False
    If pointerDisable3D = 0 And Me.Visible = False Then Me.Visible = True
    StartEventHook
    abort3Dxinput False
    xinputToDesktop = True
    RenderEnabled = True
    xenable = True
    tmrAutoRefreshIcons.Enabled = True
    tmrload.Enabled = True
    isloaded = True
End Sub



Private Sub tmrload_Timer()
    If pointerDisable3D = 1 And Me.Visible = True Then Me.Visible = False
    If pointerDisable3D = 0 And Me.Visible = False Then Me.Visible = True
    tmrload.Enabled = False
End Sub

Private Sub Form_Paint()

    If pointerDisable3D = 1 And Me.Visible = True Then Me.Visible = False
    If pointerDisable3D = 0 And Me.Visible = False Then Me.Visible = True
     
    If ispainted = True Then Exit Sub
    ispainted = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    TMRPoll.Enabled = False
    TMRPollBatteries.Enabled = False
    tmrAutoRefreshIcons.Enabled = False
    WheelUnHook Me.hWnd
    StopEventHook
    spriteClass.Terminate
    mapClass.Terminate
    buttonClass.Terminate
    itemclass.ThingsTerminate
    shadowClass.Terminate
    D3DTerminate
    consoleClass.Terminate
    soundClass.Terminate
    Set itemclass = Nothing
    Set buttonClass = Nothing
    Set ProjectileClass = Nothing
    Set mapClass = Nothing
    Set consoleClass = Nothing
    Set spriteClass = Nothing
    Set shadowClass = Nothing
    Set TMRPoll = Nothing
    Set TMRPollBatteries = Nothing
    Set xinputClass = Nothing
    Set soundClass = Nothing
    Kill App.Path & "\Game0.sav"
    Kill App.Path & "\Game1.sav"
    apiExitProcess 0
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    '
End Sub
Private Sub Form_DblClick()
    itemclass.Fire
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbKeyLButton Then
        lmousedown = True
        consoleClass.Display "Desktop"
        Dim p As POINTAPI
        If apiGetCursorPos(p) <> 0 Then oldpoint = p
        xinputToDesktop = False
        RenderEnabled = True
        TerminateEXE "WindowContextMenu.exe"
        apiSetForegroundWindow Me.hWnd
        Timer1.Enabled = True
        SetWindowPos Me.hWnd, HWND_.TOPMOST, 0, 0, 0, 0, True, True
        soundClass.ButtonOn
    ElseIf Button = vbKeyRButton Then
    ElseIf Button = vbKeyMButton Then
        frmSettings.show
        SetWindowPos frmSettings.hWnd, HWND_.TOPMOST, 0, 0, 0, 0, True, True
    End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbKeyLButton Then
        lmousedown = False
    ElseIf Button = vbKeyRButton Then
        foo = True
        itemclass.Fire
    ElseIf Button = vbKeyMButton Then
    End If
End Sub


Private Sub tmrAutoRefreshIcons_Timer()
    On Error Resume Next
    If showdesktopicons = False Then
        Exit Sub
    End If
    Dim fls() As String
    Dim fls2() As String
    Dim s As Variant
    Dim txt As String
    Dim fso As Object
    Dim f As Object
    ReDim Preserve s(0)
    Dim ds As New clsDesktopShell
    Set fso = CreateObject("Scripting.FileSystemObject")
    fls = ds.GetDesktopShellPaths(ds.fGetSpecialFolder(&H10))
    If folderinview = "" Then
        fls = ds.GetDesktopShellPaths(ds.fGetSpecialFolder(&H10))
        fls2 = ds.GetDesktopShellPaths(ds.fGetSpecialFolder(&H19))
    Else
        fls = ds.GetDesktopShellPaths(folderinview)
    End If
    For Each s In fls
        Set f = fso.GetFile(s)
        txt = txt & f.DateCreated & " "
    Next
    For Each s In fls2
        Set f = fso.GetFile(s)
        txt = txt & f.DateCreated & " "
    Next
    If txt <> oldfls Then
        oldfls = txt
        RenderEnabled = True
        load3dicons
    End If
    Set fso = Nothing
    Set f = Nothing
    If tmrAutoRefreshIcons.Interval <> 2000 Then tmrAutoRefreshIcons.Interval = 2000
End Sub




'XINPUT DEVICE
Private Sub xinputClass_OnDeviceConnected()
    isconnected = True
End Sub
Private Sub xinputClass_OnDeviceDisconnected()
    isconnected = False
End Sub
'WORLD EVENTS
Private Sub itemclass_OnTake(ByVal sender As String, ByVal index As Long)
    ' consoleClass.Display sender & " " & index
End Sub
Private Sub buttonClass_OnPush(ByVal iButton As Long)
    ' consoleClass.Display "Exit"
End Sub
Private Sub buttonClass_OnStruck(ByVal iButton As Long, ByVal jProjectile As Long)
    'consoleClass.Display CStr(iButton) & " - " & CStr(jProjectile)
    If foohittest = True Then
    End If
End Sub
Private Function bytestostring(ByVal b As Long) As String
  Dim mb As Long
  bytestostring = CStr(b) & " bytes"
  If Round(b / 1000) >= 1 Then bytestostring = CStr(Round(b / 1000)) & " KB"
  If Round(b / 1000000) >= 1 Then bytestostring = CStr(Round(b / 1000000)) & " MB"
  If Round(b / 1000000000) >= 1 Then bytestostring = CStr(Round(b / 1000000000)) & " GB"
End Function
Private Sub ProjectileClass_OnStrike(ByVal iProjectile As Long, ByVal jButton As Long, ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal r As Single)
    On Error Resume Next
    If foohittest = True Then
        foohittest = False
        Dim fso As Object
        Dim oFolder As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim st As String
        st = ThingPath(jButton)
        st = fso.getfilename(st)
        '        Set oFolder = fso.getfile(st)
        '        If oFolder Is Nothing Then Set oFolder = fso.GetFolder(st)
        '
        '        st = st & " Type:" & oFolder.Type
        ' st = st & " " & GetFileInformation(ThingPath(jButton))
        
        If InStr(1, LCase(st), ".lnk") <> 0 Then
            st = st & ", " & fso.getfilename(GetTarget(ThingPath(jButton)))

        End If
        
        Dim f As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set f = fso.GetFile(ThingPath(jButton))
        st = st & ",  Type: " & f.Type & ",  " & "Size: " & bytestostring(f.Size)
 
      
        
        If Trim(st) <> "" Then consoleClass.Display st
        Set fso = Nothing
        Set f = Nothing
        '        If (Len(st) > 0) Then
        '            Dim voice As Object
        '            Set voice = CreateObject("SAPI.SpVoice")
        '            voice.Rate = 1
        '            voice.Volume = 90
        '            voice.speak st ', 1
        '            Set voice = Nothing
        '        End If
        Dim v2 As D3DVECTOR
        v2.x = x
        v2.y = y
        v2.z = z
        soundClass.SelectItem v2, 0.7, 44100 - r * 1000
    Else
        Dim sTopic     As String
        Dim sFile      As String
        Dim sParams    As String
        Dim sDirectory As String
        abort3Dxinput False
        If foo = True Then
            foo = False
            sTopic = "Open"
            sFile = App.Path & "\WindowContextMenu.exe"
            sDirectory = vbNullString
            Dim s As String
            s = ThingPath(jButton)
            sParams = Chr(34) & s & Chr(34)
            Call RunShellExecute(sTopic, sFile, sParams, sDirectory, 1)
        Else
            sTopic = "" ' "Open"
            sFile = ThingPath(jButton)
            sDirectory = vbNullString
            sParams = ""
            Call RunShellExecute(sTopic, sFile, sParams, sDirectory, 1)
        End If
    End If
End Sub
Private Sub ProjectileClass_OnMiss(ByVal iProjectile As Long, ByVal jButton As Long, ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal r As Single)
    On Error Resume Next
    If foohittest = True Then
        foohittest = False
    Else
        Dim sTopic     As String
        Dim sFile      As String
        Dim sParams    As String
        Dim sDirectory As String
        abort3Dxinput False
        If foo = True Then
            foo = False
            sTopic = ""
            sFile = App.Path & "\WindowContextMenu.exe"
            sDirectory = vbNullString
            Dim s As String
            Dim ds   As New clsDesktopShell
            s = ds.fGetSpecialFolder(&H10)
            sParams = Chr(34) & s & Chr(34)
            Call RunShellExecute(sTopic, sFile, sParams, sDirectory, 1)
        Else
            soundClass.Damage
            consoleClass.Display "Desktop"
        End If
    End If
End Sub
Private Sub mapClass_OnLoadComplete()
   ' consoleClass.Display "Loaded "
End Sub

Private Function SimulateEventDown(ByVal index As Long)
    If index = 1 Then PressActionDown
    If index = 2 Then PressContextMenuDown
    If index = 3 Then PressCancelDown
    '    If index = 4 Then PressViewDown
    If index = 5 Then PressCloseDown
    If index = 6 Then PressMinimize
    If index = 7 Then PressMaxRestore
    If index = 8 Then PressLeftStickDown
    If index = 9 Then PressRightStickDown
    If index = 10 Then PressLeftKeyDown
    If index = 11 Then PressRightKeyDown
    If index = 12 Then PressUpKeyDown
    If index = 13 Then PressDownKeyDown
    If index = 14 Then PressAppDown
End Function
Private Function SimulateEventUp(ByVal index As Long)
    If index = 1 Then ReleaseActionUp
    If index = 2 Then ReleaseContextMenuUp
    'If index = 3 Then ReleaseCancelUp
    If index = 4 Then ReleaseViewUp
    If index = 5 Then
    End If
    'If index = 6 Then Release
    'If index = 6 Then ReleaseMinimize
    'If index = 7 Then ReleaseMaxRestore
    'If index = 8 Then ReleaseRightStickUp
    'If index = 9 Then ReleaseLeftStickUp
    'If index = 10 Then ReleaseLeftKeyUp
    'If index = 11 Then ReleaseRightKeyUp
    'If index = 12 Then ReleaseUpKeyUp
    'If index = 13 Then ReleaseDownKeyUp
    If index = 14 Then ReleaseAppUp
End Function

'XINPUT BUTTONS
Private Sub xinputClass_OnButtonADown()
    SimulateEventDown CInt(keymapA)
End Sub
Private Sub xinputClass_OnButtonAUp()
    SimulateEventUp CInt(keymapA)
End Sub
Private Sub xinputClass_OnButtonStartDown()
    SimulateEventDown CInt(keymapMenu)
End Sub
Private Sub xinputClass_OnButtonStartUp()
    SimulateEventUp CInt(keymapMenu)
End Sub
Private Sub xinputClass_OnButtonBDown()
    SimulateEventDown CInt(keymapB)
End Sub
Private Sub xinputClass_OnButtonBUp()
    SimulateEventUp CInt(keymapB)
End Sub
Private Sub xinputClass_OnButtonYDown()
    SimulateEventDown CInt(keymapY)
End Sub
Private Sub xinputClass_OnButtonYUp()
    SimulateEventUp CInt(keymapY)
End Sub
Private Sub xinputClass_OnButtonXDown()
    SimulateEventDown CInt(keymapX)
End Sub
Private Sub xinputClass_OnButtonXUp()
    SimulateEventUp CInt(keymapX)
End Sub

'PERIFERAL BUTTONS
Private Sub xinputClass_OnButtonLSHDown()
    SimulateEventDown CInt(keymapLeftBumper)
End Sub
Private Sub xinputClass_OnButtonLSHUp()
    SimulateEventUp CInt(keymapLeftBumper)
End Sub
Private Sub xinputClass_OnButtonRSHDown()
    SimulateEventDown CInt(keymapRightBumper)
End Sub
Private Sub xinputClass_OnButtonRSHUp()
    SimulateEventUp CInt(keymapRightBumper)
End Sub
Private Sub xinputClass_OnButtonLSDown()
    SimulateEventDown CInt(keymapLeftStick)
End Sub
Private Sub xinputClass_OnButtonLSUp()
    SimulateEventUp CInt(keymapLeftStick)
End Sub
Private Sub xinputClass_OnButtonRSDown()
    SimulateEventDown CInt(keymapRightStick)
End Sub
Private Sub xinputClass_OnButtonRSUp()
    SimulateEventUp CInt(keymapRightStick)
End Sub

'D-PAD NAVIGATION + DIAGONALS
Private Sub xinputClass_OnButtonLeftDown()
    SimulateEventDown CInt(keymapDLeft) '
End Sub
Private Sub xinputClass_OnButtonLeftUp()
    SimulateEventUp CInt(keymapDLeft)
End Sub
Private Sub xinputClass_OnButtonRightDown()
    SimulateEventDown CInt(keymapDRight) '
End Sub
Private Sub xinputClass_OnButtonRightUp()
    SimulateEventUp CInt(keymapDRight)
End Sub
Private Sub xinputClass_OnButtonUpDown()
    SimulateEventDown CInt(keymapDUp) '
End Sub
Private Sub xinputClass_OnButtonUpUp()
    SimulateEventUp CInt(keymapDUp) '
End Sub
Private Sub xinputClass_OnButtonDownDown()
    SimulateEventDown CInt(keymapDDown) '
End Sub
Private Sub xinputClass_OnButtonDownUp()
    SimulateEventUp CInt(keymapDDown) '
End Sub
Private Sub xinputClass_OnButtonBackDown()
    SimulateEventDown CInt(keymapChange)
End Sub
Private Sub xinputClass_OnButtonBackUp()
    SimulateEventUp CInt(keymapChange)
End Sub

Private Sub xinputClass_OnButtonNEDown()
    '
End Sub
Private Sub xinputClass_OnButtonNEUp()
    '
End Sub
Private Sub xinputClass_OnButtonSEDown()
    '
End Sub
Private Sub xinputClass_OnButtonSEUp()
    '
End Sub
Private Sub xinputClass_OnButtonNWDown()
    '
End Sub
Private Sub xinputClass_OnButtonNWUp()
    '
End Sub
Private Sub xinputClass_OnButtonSWDown()
    '
End Sub
Private Sub xinputClass_OnButtonSWUp()
    '
End Sub

'THUMBSTICKS
Private Sub xinputClass_OnLThumbChange(ByVal x As Double, ByVal y As Double)
    On Error Resume Next
    ChangeLeftThumb x, y
End Sub
Private Sub xinputClass_OnLThumbDead(ByVal x As Double, ByVal y As Double)
    On Error Resume Next
    LeftThumbDead = True
    If xinputToDesktop = True And pointerDisable2D = 0 Then
         GetWindowUnderPointer
    Else
        
    End If
End Sub
Private Sub xinputClass_OnLThumbAlive(ByVal x As Double, ByVal y As Double)
    On Error Resume Next
    LeftThumbDead = False
    If xinputToDesktop = True And pointerDisable2D = 0 Then
        RenderEnabled = False
    ElseIf xinputToDesktop = False And pointerDisable3D = 0 Then
        RenderEnabled = True
        CenterCursor
        apiSetCursorPos ScrCenterX, ScrCenterY
    Else
        RenderEnabled = False
    End If
End Sub
Private Sub xinputClass_OnRThumbChange(ByVal x As Double, ByVal y As Double)
    On Error Resume Next
    ChangeRightThumb x, y
End Sub

Private Sub xinputClass_OnRThumbDead(ByVal x As Double, ByVal y As Double)
    On Error Resume Next
    RightThumbDead = True
    If xinputToDesktop = True And pointerDisable2D = 0 Then
    
    ElseIf xinputToDesktop = False And pointerDisable3D = 0 Then
        If LeftThumbDead = True Then
            foohittest = True
            itemclass.Fire
        End If
    Else
    
    End If
End Sub
Private Sub xinputClass_OnRThumbAlive(ByVal x As Double, ByVal y As Double)
    On Error Resume Next
    RightThumbDead = False
    If xinputToDesktop = True And pointerDisable2D = 0 Then
        RenderEnabled = False
    ElseIf xinputToDesktop = False And pointerDisable3D = 0 Then
        RenderEnabled = True
        CenterCursor
        apiSetCursorPos ScrCenterX, ScrCenterY
        MSGRender = False
    Else
       RenderEnabled = False
    End If
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    SetWindowPos Me.hWnd, HWND_.TOPMOST, 0, 0, 0, 0, True, True
    Timer1.Enabled = False
End Sub
'TRIGGERS
Private Sub xinputClass_OnRTriggerChange(ByVal z As Long)
    On Error Resume Next
    ChangeRightTrigger z
End Sub
Private Sub xinputClass_OnRTriggerDown(ByVal z As Long)
    On Error Resume Next
    If xenable = True Then Exit Sub
    If xinputToDesktop = True And pointerDisable2D = 0 Then
        If LeftThumbDead = True And RightThumbDead = True Then
            pointerAcc = True
            apimouse_event MOUSEEVENTF_WHEEL, 0, 0, -1, apiGetMessageExtraInfo
        Else
            pointerAcc = False
            apimouse_event MOUSEEVENTF_.Move, 0, 0, 0, apiGetMessageExtraInfo
        End If
    ElseIf xinputToDesktop = False And pointerDisable3D = 0 Then
    Else
    End If
End Sub

Private Sub xinputClass_OnLTriggerChange(ByVal z As Long)
    On Error Resume Next
    ChangeLeftTrigger z
End Sub
Private Sub xinputClass_OnLTriggerDown(ByVal z As Long)
    On Error Resume Next
    If xenable = True Then Exit Sub
    If xinputToDesktop = True And pointerDisable2D = 0 Then
        If LeftThumbDead = True And RightThumbDead = True Then
            pointerAcc = True
            apimouse_event MOUSEEVENTF_WHEEL, 0, 0, 1, apiGetMessageExtraInfo
        Else
            pointerAcc = False
            apimouse_event MOUSEEVENTF_.Move, 0, 0, 0, apiGetMessageExtraInfo
        End If
    ElseIf xinputToDesktop = False And pointerDisable3D = 0 Then
    Else
    End If
End Sub

'DEFAULT EVENTS
Private Sub PressActionDown()
    On Error Resume Next
    If xenable = True Then Exit Sub
    Isxinput
    If IsXINPUTSupported = True And modifierkeydown = False Then Exit Sub  'ensure that we are not clicking on a taskbar window internally supported by XINPUT
    If IsAppFrameWindowForeground() = True And modifierkeydown = False Then Exit Sub     'ensure that we are not clicking on a app while it is internally supported by XINPUT
    If xinputToDesktop = True And pointerDisable2D = 0 Then
       leftmousedown = True
       apimouse_event MOUSEEVENTF_.LeftDown, 0, 0, 0, XINPUT_EXTRA_INFO
    ElseIf xinputToDesktop = False And pointerDisable3D = 0 Then
       itemclass.Fire
    Else
    End If
End Sub

Private Sub ReleaseActionUp()
    On Error Resume Next
    If leftmousedown = True Then
       leftmousedown = False
       apimouse_event MOUSEEVENTF_.LeftUp, 0, 0, 0, XINPUT_EXTRA_INFO
    End If
    If xinputToDesktop = True And pointerDisable2D = 0 Then
    ElseIf xinputToDesktop = False And pointerDisable3D = 0 Then
    Else
    End If
End Sub

Private Sub PressCancelDown()
    On Error Resume Next
    If xinputToDesktop = True And pointerDisable2D = 0 Then
        TerminateEXE "WindowContextMenu.exe"
        EscapeDialogMenuWindow
    ElseIf xinputToDesktop = False And pointerDisable3D = 0 Then
        TerminateEXE "WindowContextMenu.exe"
    Else
    End If
End Sub

' Call apiSetWindowLong(Me.hWnd, GWL_EXSTYLE, apiGetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_NOACTIVATE)
Private Sub ReleaseViewUp()
    On Error Resume Next
    If modifierkeydown = True Then
        If showdesktopicons = True Then
            showdesktopicons = False
            ThingCnt = 0
            ReDim Preserve tex(ThingCnt)
            ReDim Thing(ThingCnt)
            ReDim ThingPath(ThingCnt)
        End If
    Else
        If showdesktopicons = False Then
            showdesktopicons = True
            oldfls = ""
        End If
    End If
    If xinputToDesktop = True And pointerDisable3D = 0 Then 'Switch from 2D pointer mode -to- 3D POV mode if not disabled
        xinputToDesktop = False
        apiSetForegroundWindow Me.hWnd
        CenterCursor
        apiSetCursorPos ScrCenterX, ScrCenterY
        RenderEnabled = True
        Timer1.Enabled = True
        soundClass.ButtonOn
    ElseIf xinputToDesktop = False And pointerDisable2D = 0 Then 'Switch from POV to 2D pointer mode if not disabled
        abort3Dxinput False
        RenderEnabled = False
        SetWindowPos Me.hWnd, HWND_.NOTOPMOST, 0, 0, 0, 0, True
        SetWindowPos Me.hWnd, HWND_.bottom, 0, 0, 0, 0, False
        If frmSettings.Visible = True Then SetWindowPos frmSettings.hWnd, HWND_.TOPMOST, 0, 0, 0, 0, False, False
        If frmGamepad.Visible = True Then SetWindowPos frmGamepad.hWnd, HWND_.TOPMOST, 0, 0, 0, 0, False, False
        If frmPointer.Visible = True Then SetWindowPos frmPointer.hWnd, HWND_.TOPMOST, 0, 0, 0, 0, False, False
        If frmDisplay.Visible = True Then SetWindowPos frmDisplay.hWnd, HWND_.TOPMOST, 0, 0, 0, 0, False, False
        If frmThumb.Visible = True Then SetWindowPos frmThumb.hWnd, HWND_.TOPMOST, 0, 0, 0, 0, False, False
        soundClass.ButtonFail
     
    Else
        soundClass.Damage 'let user know that both modes are disabled currently
        If pointerDisable3D = 1 Then Me.Visible = False

    End If
       
End Sub
'        For i = 0 To UBound(minwinds)
'            If apiIsWindowVisible(minwinds(i)) <> 0 Then
'                If apiIsZoomed(minwinds(i)) <> 0 Then
'                    If IsMinimizable(minwinds(i)) = True Then apiShowWindow minwinds(i), SW_SHOWMINIMIZED
'                ElseIf apiIsIconic(minwinds(i)) <> 0 Then
'                ElseIf minwinds(i) <> Me.hWnd Then
'                    If IsMinimizable(minwinds(i)) = True Then apiShowWindow minwinds(i), SW_SHOWMINIMIZED
'                End If
'            End If
'        Next
'        xinputToDesktop = False
'        RenderEnabled = True
'        Timer1.Enabled = True
'        For i = 0 To UBound(minwinds)
'            If apiIsWindow(minwinds(i)) <> 0 Then
'                If apiIsWindowVisible(minwinds(i)) <> 0 Then
'                    If apiIsIconic(minwinds(i)) <> 0 Then
'                        wn = GetWinName(minwinds(i))
'                        If wn <> "Groove Music" And wn <> "Settings" And wn <> "Calculator" And wn <> "Microsoft Store" Then
'                            apiShowWindow minwinds(i), SW_RESTORE
'                        End If
'                    End If
'                End If
'            End If
'        Next

Private Sub PressCloseDown()
    On Error Resume Next
    If xinputToDesktop = True And pointerDisable2D = 0 Then
        If modifierkeydown = True Then
            TerminateWindowProcessUnderPointer
        Else
            CloseWindowUnderPointer
        End If
    ElseIf xinputToDesktop = False And pointerDisable3D = 0 Then
    Else
    End If
End Sub


Private Sub PressContextMenuDown()
    On Error Resume Next
    If xinputToDesktop = True And pointerDisable2D = 0 Then
        TerminateEXE "WindowContextMenu.exe"
    ElseIf xinputToDesktop = False And pointerDisable3D = 0 Then
        TerminateEXE "WindowContextMenu.exe"
    Else
    End If
End Sub
Private Sub ReleaseContextMenuUp()
    On Error Resume Next
    If xinputToDesktop = True And pointerDisable2D = 0 Then
        If IsXINPUTSupported = True Then
            apikeybd_event Keys.vk_Apps, apiOemKeyScan(Keys.vk_Apps) And &HFF, KEYEVENTF_KEYDOWN, XINPUT_EXTRA_INFO: apikeybd_event Keys.vk_Apps, apiOemKeyScan(Keys.vk_Apps) And &HFF, KEYEVENTF_KEYUP, XINPUT_EXTRA_INFO
        Else
            If modifierkeydown = True Then
                apimouse_event MOUSEEVENTF_.MiddleDown, 0, 0, 0, XINPUT_EXTRA_INFO
                apimouse_event MOUSEEVENTF_.middleUp, 0, 0, 0, XINPUT_EXTRA_INFO
            Else
                apimouse_event MOUSEEVENTF_.RightDown, 0, 0, 0, XINPUT_EXTRA_INFO
                apimouse_event MOUSEEVENTF_.RightUp, 0, 0, 0, XINPUT_EXTRA_INFO
            End If
        End If
    ElseIf xinputToDesktop = False And pointerDisable3D = 0 Then
        foo = True
        itemclass.Fire
    Else
    End If
End Sub
Private Sub PressAppDown()
    modifierkeydown = True
    modifieractive = False
End Sub

Private Sub ReleaseAppUp()
    On Error Resume Next
    modifierkeydown = False
    If modifieractive = False Then
        abort3Dxinput False
        RenderEnabled = False
        If pointerDisable3D = 0 Then SetWindowPos Me.hWnd, HWND_.NOTOPMOST, 0, 0, 0, 0, True
        If pointerDisable3D = 0 Then SetWindowPos Me.hWnd, HWND_.bottom, 0, 0, 0, 0, False
        
        frmSettings.show
        SetWindowPos frmSettings.hWnd, HWND_.TOPMOST, 0, 0, 0, 0, False, False
    End If
End Sub
Private Sub ChangeLeftThumb(ByVal x As Double, ByVal y As Double)
    On Error Resume Next
    If frmThumb.Visible = True Then
        If x > 0 Then frmThumb.Label6.Caption = x Else frmThumb.Label6.Caption = 0
        If x < 0 Then frmThumb.Label10.Caption = x Else frmThumb.Label10.Caption = 0
        If y < 0 Then frmThumb.Label8.Caption = y Else frmThumb.Label8.Caption = 0
        If y > 0 Then frmThumb.Label3.Caption = y Else frmThumb.Label3.Caption = 0
    End If
    Dim leftvector As Vector2
    leftvector.x = CDbl(x / 32767)
    leftvector.y = CDbl(y / 32767)
    oldlv = leftvector
    oldis.gamepad.sThumbLX = x
    oldis.gamepad.sThumbLY = y
End Sub
Private Sub ChangeRightThumb(ByVal x As Double, ByVal y As Double)
    On Error Resume Next
    If frmThumb.Visible = True Then
        frmThumb.Label7.Caption = x
        frmThumb.Label9.Caption = y
        If x > 0 Then frmThumb.Label7.Caption = x Else frmThumb.Label7.Caption = 0
        If x < 0 Then frmThumb.Label11.Caption = x Else frmThumb.Label11.Caption = 0
        If y < 0 Then frmThumb.Label9.Caption = y Else frmThumb.Label9.Caption = 0
        If y > 0 Then frmThumb.Label4.Caption = y Else frmThumb.Label4.Caption = 0
    End If
    Dim rightvector As Vector2
    rightvector.x = CDbl(x / 32767)
    rightvector.y = CDbl(y / 32767)
    oldrv = rightvector
    oldis.gamepad.sThumbRX = x
    oldis.gamepad.sThumbRY = y
End Sub
Private Sub ChangeRightTrigger(ByVal z As Long)
    oldis.gamepad.bRightTrigger = z
End Sub
Private Sub ChangeLeftTrigger(ByVal z As Long)
    oldis.gamepad.bLeftTrigger = z
End Sub
Private Sub PressUpKeyDown()
    On Error Resume Next
    If xenable = True Then Exit Sub
    Isxinput
    If IsXINPUTSupported = True Then Exit Sub
    If IsAppWindowForeground() = True Then Exit Sub
    If xinputToDesktop = True And pointerDisable2D = 0 Then
        apikeybd_event Keys.vk_up, apiOemKeyScan(Keys.vk_up) And &HFF, KEYEVENTF_KEYDOWN, XINPUT_EXTRA_INFO: apikeybd_event Keys.vk_up, apiOemKeyScan(Keys.vk_up) And &HFF, KEYEVENTF_KEYUP, XINPUT_EXTRA_INFO
    ElseIf xinputToDesktop = False And pointerDisable3D = 0 Then
    Else
    End If
End Sub
Private Sub PressDownKeyDown()
    On Error Resume Next
    If xenable = True Then Exit Sub
    Isxinput
    If IsXINPUTSupported = True Then Exit Sub
    If IsAppWindowForeground() = True Then Exit Sub
    If xinputToDesktop = True And pointerDisable2D = 0 Then
        apikeybd_event Keys.vk_down, apiOemKeyScan(Keys.vk_down) And &HFF, KEYEVENTF_KEYDOWN, XINPUT_EXTRA_INFO: apikeybd_event Keys.vk_down, apiOemKeyScan(Keys.vk_down) And &HFF, KEYEVENTF_KEYUP, XINPUT_EXTRA_INFO
    ElseIf xinputToDesktop = False And pointerDisable3D = 0 Then
    Else
    End If
End Sub
Private Sub PressLeftKeyDown()
    On Error Resume Next
    If xenable = True Then Exit Sub
    Isxinput
    If IsXINPUTSupported = True Then Exit Sub
    If IsAppWindowForeground() = True Then Exit Sub
    If xinputToDesktop = True And pointerDisable2D = 0 Then
        apikeybd_event Keys.vk_Left, apiOemKeyScan(Keys.vk_Left) And &HFF, KEYEVENTF_KEYDOWN, XINPUT_EXTRA_INFO: apikeybd_event Keys.vk_Left, apiOemKeyScan(Keys.vk_Left) And &HFF, KEYEVENTF_KEYUP, XINPUT_EXTRA_INFO
    ElseIf xinputToDesktop = False And pointerDisable3D = 0 Then
    Else
    End If
End Sub
Private Sub PressRightKeyDown()
    On Error Resume Next
    If xenable = True Then Exit Sub
    Isxinput
    If IsXINPUTSupported = True Then Exit Sub
    If IsAppWindowForeground() = True Then Exit Sub
    If xinputToDesktop = True And pointerDisable2D = 0 Then
        apikeybd_event Keys.vk_Right, apiOemKeyScan(Keys.vk_Right) And &HFF, KEYEVENTF_KEYDOWN, XINPUT_EXTRA_INFO: apikeybd_event Keys.vk_Right, apiOemKeyScan(Keys.vk_Right) And &HFF, KEYEVENTF_KEYUP, XINPUT_EXTRA_INFO
    ElseIf xinputToDesktop = False And pointerDisable3D = 0 Then
    Else
    End If
End Sub

Private Sub EscapeDialogMenuWindow()
    On Error Resume Next
    If pointerDisable2D = 1 Then Exit Sub
    Dim fwnd As Long
    fwnd = apiGetForegroundWindow
    If apiIsWindow(fwnd) <> 0 Then
        Dim wn As WINNAME
        wn = GetWinNameAPI(fwnd, False, True)
        If wn.lpClass = "" Or wn.lpClass = "Windows.UI.Core.CoreWindow" Then
            wn = GetWinNameAPI(fwnd, True, False)
            If Trim(LCase(wn.lpText)) = "action center" Or Trim(LCase(wn.lpText)) = "windows shell experience host" Then
                winescape
                Exit Sub
            End If
        End If
    End If
    If xenable = True Then Exit Sub
    apikeybd_event Keys.vk_Escape, apiOemKeyScan(Keys.vk_Escape) And &HFF, KEYEVENTF_KEYDOWN, XINPUT_EXTRA_INFO: apikeybd_event Keys.vk_Escape, apiOemKeyScan(Keys.vk_Escape) And &HFF, KEYEVENTF_KEYUP, XINPUT_EXTRA_INFO
    
    fwnd = apiGetForegroundWindow
    
    If apiIsWindow(fwnd) = 0 Or fwnd = apiFindWindow("Shell_TrayWnd", vbNullString) Or fwnd = apiFindWindow("Progman", "Program Manager") Then
        apikeybd_event Keys.VK_LWIN, apiOemKeyScan(Keys.VK_LWIN) And &HFF, KEYEVENTF_KEYDOWN, XINPUT_EXTRA_INFO
        apikeybd_event Keys.vk_b, apiOemKeyScan(Keys.vk_b) And &HFF, KEYEVENTF_KEYDOWN, XINPUT_EXTRA_INFO: apikeybd_event Keys.vk_b, apiOemKeyScan(Keys.vk_b) And &HFF, KEYEVENTF_KEYUP, XINPUT_EXTRA_INFO
        apikeybd_event Keys.VK_LWIN, apiOemKeyScan(Keys.VK_LWIN) And &HFF, KEYEVENTF_KEYUP, XINPUT_EXTRA_INFO
    End If
End Sub

Private Function GetWindowUnderPointer() As Long
    On Error Resume Next
    Dim p As POINTAPI
    Dim hWnd As Long
    If apiGetCursorPos(p) <> 0 Then hWnd = GetForegroundFromPoint(p)
    If hWnd = 0 Then Exit Function
    If apiIsWindow(hWnd) = 0 Then Exit Function
    If hWnd = apiGetDesktopWindow Then Exit Function
    If hWnd = apiFindWindow("Progman", "Program Manager") Then Exit Function
    If hWnd = apiFindWindow("Shell_TrayWnd", vbNullString) Then Exit Function
    'If hWnd = apiFindWindow("WorkerW", vbNullString) Then Exit Function
    If hWnd = apiFindWindow("ThunderRT6FormDC", "Window Launcher") Then Exit Function
    If hWnd = apiFindWindow("ThunderRT6FormDC", "Window Wallpaper") Then Exit Function
    If hWnd = apiFindWindow("ThunderRT6FormDC", "Start Menu") Then Exit Function
    If hWnd = Me.hWnd Then Exit Function
    pullwnd = hWnd
    GetWindowUnderPointer = hWnd
End Function

Private Sub PressLeftStickDown()
    On Error Resume Next
    If pointerDisable2D = 1 Then Exit Sub
    If xenable = True Then Exit Sub
    If IsXINPUTSupported = True Then Exit Sub
    Dim hWnd As Long
    hWnd = GetWindowUnderPointer
    If hWnd = 0 Then Exit Sub
    If apiIsWindow(hWnd) = 0 Then Exit Sub
    '    If hWnd = apiFindWindow("ThunderRT6FormDC", "Window Launcher") Then
    '        SetWindowPos hWnd, HWND_.NOTOPMOST, 0, 0, 0, 0, True
    '        Exit Sub
    '    End If
    SetWindowPos hWnd, HWND_.NOTOPMOST, 0, 0, 0, 0, True
    SetWindowPos hWnd, HWND_.bottom, 0, 0, 0, 0, False
    SetWindowPos Me.hWnd, HWND_.bottom, 0, 0, 0, 0, False
    GetWindowUnderPointer
End Sub

Private Sub PressRightStickDown()
    On Error Resume Next
      If pointerDisable2D = 1 Then Exit Sub
    If xenable = True Then Exit Sub
    If IsXINPUTSupported = True Then Exit Sub
    Dim hWnd As Long
    hWnd = GetWindowUnderPointer
    If hWnd = 0 Then Exit Sub
    If apiIsWindow(hWnd) = 0 Then Exit Sub
    '    If hWnd = apiFindWindow("ThunderRT6FormDC", "Window Launcher") Then
    '        SetWindowPos apiFindWindow("ThunderRT6FormDC", "Window Launcher"), HWND_.TOPMOST, 0, 0, 0, 0, True
    '        Exit Sub
    '    End If
    SetWindowPos hWnd, HWND_.TOPMOST, 0, 0, 0, 0, True
    SetWindowPos hWnd, HWND_.NOTOPMOST, 0, 0, 0, 0, True
    apiSetForegroundWindow hWnd
    GetWindowUnderPointer
End Sub
Private Sub PressMaxRestore()
    On Error Resume Next
      If pointerDisable2D = 1 Then Exit Sub
    Dim hWnd As Long
    hWnd = GetWindowUnderPointer
    If hWnd = 0 Then Exit Sub
    If apiIsWindow(hWnd) = 0 Then Exit Sub
    If apiIsZoomed(hWnd) <> 0 Then
        apiShowWindow hWnd, SW_RESTORE
    ElseIf apiIsIconic(hWnd) <> 0 Then
        apiShowWindow hWnd, SW_RESTORE
    ElseIf hWnd <> Me.hWnd Then
        If IsMaximizable(hWnd) = True Then
            apiShowWindow hWnd, SW_SHOWMAXIMIZED
        End If
    End If
    GetWindowUnderPointer
End Sub
Private Sub PressMinimize()
    On Error Resume Next
      If pointerDisable2D = 1 Then Exit Sub
    Dim hWnd As Long
    hWnd = GetWindowUnderPointer
    If hWnd = 0 Then Exit Sub
    If apiIsWindow(hWnd) = 0 Then Exit Sub
    If IsMinimizable(hWnd) = True Then
        apiShowWindow hWnd, SW_SHOWMINIMIZED
        GetWindowUnderPointer
    End If
End Sub

Private Sub TMRPoll_Timer()
    On Error Resume Next
    Dim r  As Byte
    Dim l  As Byte
    Dim d  As Double
    Dim d2 As Double
    r = oldis.gamepad.bRightTrigger
    l = oldis.gamepad.bLeftTrigger
    d = 1
    d2 = 1
    If r > 0 Then d = (r / 255) * pointerMaxPointerAcceleration
    If r > 0 Then d2 = (r / 255) * pointerMaxPOVAcceleration
    If d < 1 Then d = 1
    If d2 < 1 Then d2 = 1
    Dim v As Vector2
    Dim x As Double
    Dim y As Double
    Dim xy As Double
    xy = 0.001
    v = oldlv
    x = Abs(v.x)
    y = Abs(v.y)
    Dim thumbThrottle As Double
    Dim thumbThrottle2 As Double
    If LeftThumbDead = True Then
        If xinpleft = True Then xinpleft = False
        If xinpright = True Then xinpright = False
        If xinpforward = True Then xinpforward = False
        If xinpbackward = True Then xinpbackward = False
    Else
        If isconnected = True Then
            Dim z             As Double
            z = (x ^ 2) + (y ^ 2)
            thumbThrottle = z ^ (0.5)
            If thumbThrottle > 1 Then thumbThrottle = 1
            If xinputToDesktop = True And pointerDisable2D = 0 Then
                thumbThrottle = thumbThrottle + pointerMaxPointerSpeed
                Isxinput
                Dim Offset        As New Point
                If d < 1 Then d = 1
                thumbThrottle = thumbThrottle + d - ((l / 255) * pointerMaxPointerAcceleration)
                If thumbThrottle < 1 Then thumbThrottle = 1
                Offset.x = Round(v.x * thumbThrottle)
                Offset.y = Round(v.y * thumbThrottle)
                If IsXINPUTSupported = False Or modifierkeydown = True Then
                    apimouse_event MOUSEEVENTF_.Move, Offset.x, -Offset.y, 0, apiGetMessageExtraInfo
                End If
            ElseIf xinputToDesktop = False And pointerDisable3D = 0 Then
                thumbThrottle = thumbThrottle + pointerMaxPOVSpeed
                If d2 < 1 Then d2 = 1
                thumbThrottle = thumbThrottle + d2
                If v.y > 0.1 Then
                    xy = Abs(v.y)
                    xinpbackward = False
                    xinpforward = True
                ElseIf v.y < -0.1 Then
                    xy = Abs(v.y)
                    xinpforward = False
                    xinpbackward = True
                Else
                    xinpforward = False
                    xinpbackward = False
                End If
                If v.x > 0.5 Then
                    xy = Abs(v.x)
                    xinpleft = False
                    xinpright = True
                ElseIf v.x < -0.5 Then
                    xy = Abs(v.x)
                    xinpright = False
                    xinpleft = True
                Else
                    xinpleft = False
                    xinpright = False
                End If
            Else
            End If
        End If
    End If
    Dim v2             As Vector2
    Dim X2             As Double
    Dim Y2             As Double
    Dim b              As Boolean
    Dim Offset2        As POINTAPI
    v2 = oldrv
    X2 = Abs(v2.x)
    Y2 = Abs(v2.y)
    If RightThumbDead = True Then
    Else
        If isconnected = True Then
            Dim z2 As Double
            z2 = (X2 ^ 2) + (Y2 ^ 2)
            thumbThrottle2 = z2 ^ (0.5)
            If thumbThrottle2 > 1 Then thumbThrottle2 = 1
            If xinputToDesktop = True Then
                thumbThrottle2 = thumbThrottle2 + pointerMaxPointerSpeed
                If d < 1 Then d = 1
                thumbThrottle2 = thumbThrottle2 + d - ((l / 255) * pointerMaxPointerAcceleration)
            Else
                thumbThrottle2 = thumbThrottle2 + pointerMaxPOVSpeed
                If d2 < 1 Then d2 = 1
                thumbThrottle2 = thumbThrottle2 + d2
            End If
            If thumbThrottle2 < 4 Then thumbThrottle2 = 4
            Offset2.x = (v2.x * thumbThrottle2)
            Offset2.y = (v2.y * thumbThrottle2)
            If xinputToDesktop = True And pointerDisable2D = 0 Then
                If apiIsWindow(pullwnd) = 0 Then pullwnd = GetWindowUnderPointer
                If apiIsWindow(pullwnd) <> 0 Then
                    If apiIsZoomed(pullwnd) = 0 And apiIsIconic(pullwnd) = 0 Then
                        Dim rct As RECT
                        If apiGetWindowRect(pullwnd, rct) <> 0 Then
                            If modifierkeydown = True Then
                                If IsSizable(pullwnd) = True Then
                                    Dim w As Long
                                    Dim h As Long
                                    w = (rct.right - rct.left) + Offset2.x
                                    h = (rct.bottom - rct.top) - Offset2.y
                                    If w > 16 And h > 16 Then
                                        apiMoveWindow pullwnd, rct.left, rct.top, w, h, True
                                    End If
                                End If
                            Else
                                Dim p As New Point
                                p.x = rct.left
                                p.y = rct.top
                                Dim selp As POINTAPI
                                selp.x = p.x + Offset2.x
                                selp.y = p.y - Offset2.y
                                If selp.x > (VirtualScreenWidth - 100) Then
                                    selp.x = (VirtualScreenWidth - 100)
                                End If
                                If selp.y > (VirtualScreenHeight - 100) Then
                                    selp.y = (VirtualScreenHeight - 100)
                                End If
                                If selp.x < (-(rct.right - rct.left) + 100) Then
                                    selp.x = (-(rct.right - rct.left) + 100)
                                End If
                                If selp.y < (-(rct.bottom - rct.top) + 100) Then
                                    selp.y = (-(rct.bottom - rct.top) + 100)
                                End If
                                apiMoveWindow pullwnd, selp.x, selp.y, rct.right - rct.left, rct.bottom - rct.top, True
                            End If
                        End If
                    End If
                End If
            ElseIf xinputToDesktop = False And pointerDisable3D = 0 Then
            Else
            End If
        End If
    End If
    Dim ts       As D3DVECTOR
    Dim mx       As Single
    Dim my_      As Single
    Static fJump As Boolean
    If lmousedown = True Then
        Dim pC As POINTAPI
        If apiGetCursorPos(pC) <> 0 Then
            mx = (pC.x - oldpoint.x) * MouseSensX * pointerMaxPOVSpeed
            my_ = (pC.y - oldpoint.y) * MouseSensY * pointerMaxPOVSpeed
            oldpoint = pC
        End If
    Else
        If xinputToDesktop = False Then
            mx = -Offset2.x / 200
            my_ = (Offset2.y / 200)
        End If
    End If
    PlAngle = PlAngle + -mx
    PlDiff = PlDiff + -my_
    If PlAngle < 0 Then
        PlAngle = PlAngle + 2 * Pi
    ElseIf PlAngle > 2 * Pi Then
        PlAngle = PlAngle - 2 * Pi
    End If
    If PlDiff < -1.55 Then
        PlDiff = -1.55
    ElseIf PlDiff > 1.55 Then
        PlDiff = 1.55
    End If
    SinA = Sin(PlAngle)
    CosA = Cos(PlAngle)
    SinD = Sin(PlDiff)
    CosD = Cos(PlDiff)
    If xinputToDesktop = False And pointerDisable3D = 0 Then
        ts = Vec3(0, 0, 0)
        If xinpforward = True Then
            ts.z = (ts.z + CosA)
            ts.x = (ts.x + SinA)
        ElseIf xinpbackward = True Then
            ts.z = (ts.z - CosA) '
            ts.x = (ts.x - SinA)
        End If
        If xinpleft = True Then
            ts.z = (ts.z + SinA)
            ts.x = (ts.x - CosA)
        ElseIf xinpright = True Then
            ts.z = (ts.z - SinA)
            ts.x = (ts.x + CosA)
        End If
        SetTargetSpeed ts, xy
    ElseIf xinputToDesktop = True And pointerDisable2D = 0 Then
        TriggerLeftTick
        TriggerRightTick
    End If
    If RenderEnabled = True Then
        NowTime = QTime
        If NowTime > OldTime + (directxQuant * 100) Then OldTime = NowTime - (directxQuant * 100)
        Do
            NowTime = QTime
            If NowTime <= (OldTime + directxQuant) Then Exit Do
            OldTime = OldTime + directxQuant
            buttonClass.Tick
            PhysTick
            SetCamera
            itemclass.SelectorTick
            itemclass.Tick
            consoleClass.Tick
        Loop
        Select Case Dev.TestCooperativeLevel
            Case D3DERR_DEVICELOST
                Exit Sub
            Case D3DERR_DEVICENOTRESET
                ResetStates
        End Select
        shadowClass.DrawBegin
        mapClass.DrawShade
        LS.DrawShade
        itemclass.DrawShade
        shadowClass.DrawEnd
        Dev.Clear D3DCLEAR_ZBUFFER, &H0, 1, 0
        If Dev.BeginScene Then
            Dev.SetPixelShaderConstantF 0, VarPtr(PSConst), 3
            mapClass.Draw
            itemclass.Draw
            consoleClass.Draw
            Dev.EndScene
        End If
        Dev.Present
        apiRedrawWindow Me.hWnd, False, 0, RDW_UPDATENOW Or RDW_NOERASE Or RDW_NOINTERNALPAINT 'Redraw this window (invoke a Paint-event)  RDW_NOINTERNALPAINT
    End If
End Sub
Private Sub TriggerLeftTick()
    On Error Resume Next
    If pointerDisable2D = 1 Then Exit Sub
    If xenable = True Then Exit Sub
    If pointerAcc = False Then Exit Sub
    If LeftThumbDead = False Or RightThumbDead = False Then Exit Sub
    Dim tl As Long
    Dim tr As Long
    Dim v  As Long
    Dim wd As Long
    tl = oldis.gamepad.bLeftTrigger
    tr = oldis.gamepad.bRightTrigger
    v = CInt(((tl) / 255) * 100)
    tl = tl - tr 'actively braking
    If tl < 1 Then Exit Sub 'yield to opposite
    wd = tl
    If wd <> 0 And wd <> 255 Then rollingscroll = 0
    If wd = 255 Then
        wd = CInt((wd + rollingscroll) * 2)
        rollingscroll = rollingscroll + pointerMaxWheelAcceleration
    End If
    wd = CInt(wd / 200)
    wd = wd * pointerMaxWheelSpeed
    If modifierkeydown = True Then
        apimouse_event MOUSEEVENTF_HWHEEL, 0, 0, -wd, apiGetMessageExtraInfo
    Else
        apimouse_event MOUSEEVENTF_WHEEL, 0, 0, wd, apiGetMessageExtraInfo
    End If
End Sub
Private Sub TriggerRightTick()
    On Error Resume Next
     If pointerDisable2D = 1 Then Exit Sub
    If xenable = True Then Exit Sub
    If pointerAcc = False Then Exit Sub
    If LeftThumbDead = False Or RightThumbDead = False Then Exit Sub
    Dim tr As Long
    Dim tl As Long
    Dim v  As Long
    Dim wd As Long
    tr = oldis.gamepad.bRightTrigger
    tl = oldis.gamepad.bLeftTrigger
    v = CInt(((tr) / 255) * 100)
    tr = tr - tl 'actively braking
    If tr < 1 Then Exit Sub 'yield to opposite
    wd = tr
    If wd <> 0 And wd <> 255 Then rollingscroll = 0
    If wd = 255 Then
        wd = CInt((wd + rollingscroll) * 2)
        rollingscroll = rollingscroll + pointerMaxWheelAcceleration
    End If
    wd = CInt(wd / 200)
    wd = wd * pointerMaxWheelSpeed
    If modifierkeydown = True Then
        apimouse_event MOUSEEVENTF_HWHEEL, 0, 0, wd, apiGetMessageExtraInfo
    Else
        apimouse_event MOUSEEVENTF_WHEEL, 0, 0, -wd, apiGetMessageExtraInfo
    End If
End Sub
Private Sub TMRPollBatteries_Timer()
    On Error Resume Next
    Dim batlev As String
    batlev = xinputClass.BatteryLevel(active_input_user)
End Sub

'Private Sub tmrWheel1_Timer()
'    On Error Resume Next
'    If xinputToDesktop = False Then
'        If scrollposi = 0 Then
'            If xinpbackward = True Then xinpbackward = False
'            If xinpforward = True Then xinpforward = False
'        ElseIf scrollposi = Abs(scrollposi) Then
'            If xinpbackward = True Then xinpbackward = False
'            If xinpforward = False Then xinpforward = True
'        ElseIf scrollposi = -Abs(scrollposi) Then
'            If xinpforward = True Then xinpforward = False
'            If xinpbackward = False Then xinpbackward = True
'        End If
'    End If
'End Sub

Private Sub winescape()
    On Error Resume Next
    apikeybd_event Keys.VK_LWIN, apiOemKeyScan(Keys.VK_LWIN) And &HFF, KEYEVENTF_KEYDOWN, XINPUT_EXTRA_INFO
    apikeybd_event Keys.vk_b, apiOemKeyScan(Keys.vk_b) And &HFF, KEYEVENTF_KEYDOWN, XINPUT_EXTRA_INFO: apikeybd_event Keys.vk_b, apiOemKeyScan(Keys.vk_b) And &HFF, KEYEVENTF_KEYUP, XINPUT_EXTRA_INFO
    apikeybd_event Keys.VK_LWIN, apiOemKeyScan(Keys.VK_LWIN) And &HFF, KEYEVENTF_KEYUP, XINPUT_EXTRA_INFO
End Sub
Private Sub disconnectedgamepad()
    LeftThumbDead = True
    RightThumbDead = True
    isconnected = False
    PhysInit
    oldlv.x = 0
    oldlv.y = 0
    oldrv.x = 0
    oldrv.y = 0
End Sub
Private Sub CloseWindowUnderPointer()
    On Error Resume Next
      If pointerDisable2D = 1 Then Exit Sub
    If xenable = True Then Exit Sub
    Dim p As POINTAPI
    If apiGetCursorPos(p) = 0 Then Exit Sub
    Dim pwnd As Long
    pwnd = GetForegroundFromPoint(p)
    If pwnd = 0 Or apiIsWindow(pwnd) = 0 Then Exit Sub
    If pwnd = Me.hWnd Then Exit Sub
    If pwnd = apiFindWindow("Progman", "Program Manager") Then Exit Sub
    If pwnd = apiFindWindow("Shell_TrayWnd", "") Then Exit Sub
    If pwnd = apiGetDesktopWindow Then Exit Sub
    If pwnd = apiFindWindow("ThunderRT6FormDC", "Window Launcher") Then Exit Sub
    If pwnd = apiFindWindow("ThunderRT6FormDC", "Window Wallpaper") Then Exit Sub
    If pwnd = apiFindWindow("ThunderRT6FormDC", "Start Menu") Then Exit Sub 'apiShowWindow pwnd, SW_HIDE
    cwnd = pwnd
    closewin
    GetWindowUnderPointer
End Sub
Private Sub TerminateWindowProcessUnderPointer()
    On Error Resume Next
    If pointerDisable2D = 1 Then Exit Sub
    If xenable = True Then Exit Sub
    Dim p As POINTAPI
    If apiGetCursorPos(p) = 0 Then Exit Sub
    Dim pwnd As Long
    pwnd = GetForegroundFromPoint(p)
    If pwnd = 0 Or apiIsWindow(pwnd) = 0 Then Exit Sub
    If pwnd = Me.hWnd Then Exit Sub
    If pwnd = apiFindWindow("Progman", "Program Manager") Then Exit Sub
    If pwnd = apiFindWindow("Shell_TrayWnd", "") Then Exit Sub
    If pwnd = apiGetDesktopWindow Then Exit Sub
    If pwnd = apiFindWindow("ThunderRT6FormDC", "Window Launcher") Then Exit Sub
    If pwnd = apiFindWindow("ThunderRT6FormDC", "Window Wallpaper") Then Exit Sub
    If pwnd = apiFindWindow("ThunderRT6FormDC", "Start Menu") Then Exit Sub 'apiShowWindow pwnd, SW_HIDE
    cwnd = pwnd
    Dim pid As Long
    apiGetWindowThreadProcessId cwnd, pid
    If pid <> 0 Then
        Dim lprocess As Long
        lprocess = apiOpenProcess(PROCESS_TERMINATE, 0, pid)
        If lprocess <> 0 Then
            apiTerminateProcess lprocess, 0
            apiCloseHandle lprocess
        End If
    Else
        closewin
    End If
    GetWindowUnderPointer
End Sub
Private Sub closewin()
    On Error Resume Next
    If apiIsWindow(cwnd) <> 0 Then
        Call apiSendMessageTimeout(cwnd, WM_CLOSE, 0, 0, SMTO_ABORTIFHUNG, 400, 0)
        cwnd = 0
    End If
End Sub

Friend Function FileExtractIcon(ByVal sFileName As String) As StdPicture
    On Error GoTo skip
    Dim oPic                       As StdPicture
    Dim tPic                       As PICTDESC
    Dim tIDispatch                 As GUID
    Dim hIcon                      As Long
    Dim tFileInfo                  As SHFILEINFO
    Const SHGFI_ICON               As Long = &H100
    Const SHGFI_DISPLAYNAME        As Long = &H200
    Const SHGFI_TYPENAME           As Long = &H400
    Const SHGFI_SMALLICON          As Long = &H1
    Const SHGFI_LARGEICON          As Long = &H0
    Const SHGFI_OPENICON           As Long = &H2         ';     // get open icon
    '  Const SHGFI_LARGEICON As Long = &H0        ';     // get large icon
    ' Const SHGFI_SMALLICON   As Long = &H1       ';     // get small icon
    Const SHIL_JUMBO               As Long = &H4       ';     // get jumbo icon 256x256
    Const SHIL_EXTRALARGE          As Long = &H2       ';     // get extra large icon 48x48
    'Const SHIL_LARGE   As Long = &H0       ';     // get large icon 32x32
    'Const SHIL_SMALL   As Long = &H1       ';     // get small icon 16x16
    'Const SHIL_SYSSMALL   As Long = &H3       ';     // get icon based off of GetSystemMetrics
    Const FILE_ATTRIBUTE_READONLY  As Long = &H1
    Const FILE_ATTRIBUTE_HIDDEN    As Long = &H2
    Const FILE_ATTRIBUTE_SYSTEM    As Long = &H4
    Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
    Const FILE_ATTRIBUTE_ARCHIVE   As Long = &H20
    Const FILE_ATTRIBUTE_NORMAL    As Long = &H80
    Const FILE_ATTRIBUTE_TEMPORARY As Long = &H100
    Call SHGetFileInfo(sFileName, 0, tFileInfo, Len(tFileInfo), SHGFI_DISPLAYNAME Or SHGFI_TYPENAME Or SHGFI_LARGEICON Or SHGFI_ICON)
    '  Call SHGetFileInfo(sFileName, 0, tFileInfo, Len(tFileInfo), SHGFI_ICON Or SHGFI_OPENICON Or SHIL_JUMBO)
    hIcon = tFileInfo.hIcon
    '    Dim fn As String
    '    fn = sFileName
    ''    Dim dot As String
    ''    dot = Mid(fn, Len(fn) - 4, 1)
    ''
    '    fn = Replace(fn, ".lnk", "")
    '    If Len(Trim(Dir(fn, vbNormal))) = 0 Then
    '        Dim hIML         As IUnknown
    '        Dim GUID(0 To 3) As Long
    '        Dim lResult      As Long
    '        Dim lIconSize    As Long
    '        lIconSize = SHIL_EXTRALARGE
    '        If IIDFromString(StrPtr(IID_IImageList), GUID(0)) = 0 Then
    '            lResult = SHGetImageList(lIconSize, GUID(0), ByVal VarPtr(hIML))
    '            If lResult = 0& Then
    '                hIcon = ImageList_GetIcon(ObjPtr(hIML), 3, 0)
    '            End If
    '        End If
    '    End If
    With tPic
        .cbSize = Len(tPic)
        .PicType = 3  'vbPicTypeIcon
        .hImage = hIcon
    End With
    With tIDispatch 'Fill IDispatch Interface ID,{00020400-0000-0000-C000-000000046}
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    Call OleCreatePictureIndirect(tPic, tIDispatch, 0, oPic)
    Set FileExtractIcon = oPic
    Exit Function
skip:
    Set FileExtractIcon = Nothing
    Set oPic = Nothing
    On Error GoTo 0
End Function
Friend Function GetFileInformation(ByVal fileFullPath As String) As String
    On Error Resume Next
    Dim lDummy              As Long
    Dim lSize               As Long
    Dim RC                  As Long
    Dim lVerbufferLen       As Long
    Dim sBuffer()           As Byte
    Dim lBufferLen          As Long
    Dim bytebuffer(260)     As Byte
    Dim Lang_Charset_String As String
    Dim HexNumber           As Long
    Dim Buffer              As String
    Dim lVerPointer         As Long
    Dim ProdName            As String
    GetFileInformation = ""
    Buffer = String(260, 0)
    lBufferLen = apiGetFileVersionInfoSize(fileFullPath, lDummy)
    If lBufferLen >= 1 Then
        ReDim sBuffer(lBufferLen)
        RC = apiGetFileVersionInfo(fileFullPath, 0&, lBufferLen, sBuffer(0))
        If RC <> 0 Then
            RC = apiVerQueryValueByteLong(sBuffer(0), "\VarFileInfo\Translation", lVerPointer, lBufferLen)
            If RC <> 0 Then
                Call apiMoveMemory(bytebuffer(0), lVerPointer, lBufferLen)
                HexNumber = bytebuffer(2) + bytebuffer(3) * &H100 + bytebuffer(0) * &H10000 + bytebuffer(1) * &H1000000
                Lang_Charset_String = Hex(HexNumber)
                Do While Len(Lang_Charset_String) < 8
                    Lang_Charset_String = "0" & Lang_Charset_String
                Loop
                GetFileInformation = GetStringValue("\StringFileInfo\" & Lang_Charset_String & "\" & "FileDescription", lVerPointer, lBufferLen, sBuffer)
                If Trim(GetFileInformation) = "" Then GetFileInformation = GetStringValue("\StringFileInfo\" & Lang_Charset_String & "\" & "ProductName", lVerPointer, lBufferLen, sBuffer)
                If Trim(GetFileInformation) = "" Then GetFileInformation = GetStringValue("\StringFileInfo\" & Lang_Charset_String & "\" & "OriginalFileName", lVerPointer, lBufferLen, sBuffer)
                If Trim(GetFileInformation) = "" Then GetFileInformation = GetStringValue("\StringFileInfo\" & Lang_Charset_String & "\" & "InternalName", lVerPointer, lBufferLen, sBuffer)
            End If
        End If
    End If
End Function
Private Sub load3dicons()
    On Error Resume Next
    If pointerDisable3D = 1 Then Exit Sub
    Dim xstart   As Long
    Dim xend     As Long
    Dim m        As Long
    Dim n        As Long
    Dim xcurrent As Long
    Dim zstart   As Long
    Dim zend     As Long
    Dim zcurrent As Long
    Dim z        As Long
    xstart = 190
    xend = 230
    xcurrent = xstart
    zstart = 0
    zend = 15
    zcurrent = zstart
    Dim ds   As New clsDesktopShell
    Dim s()  As String
    Dim s2() As String
    Dim spfo As String
    If folderinview = "" Then
        s = ds.GetDesktopShellPaths(ds.fGetSpecialFolder(&H10))
        s2 = ds.GetDesktopShellPaths(ds.fGetSpecialFolder(&H19))
    Else
        s = ds.GetDesktopShellPaths(folderinview)
    End If
    ThingCnt = 0
    ReDim Preserve tex(ThingCnt)
    ' Set tex(ThingCnt - 1) = CreateTextureFromFileEx(Dev, App.Path & "\3DEngine\Meshes\" & texturename, 0, 0, 0, D3DUSAGE_NONE, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_BOX, D3DX_FILTER_LINEAR, 0)
    ReDim Thing(ThingCnt)
    ReDim ThingPath(ThingCnt)
    Dim st   As Variant
    Dim i    As Long
    Dim finf As String
    Dim rct  As RECT
    With rct
        .left = 0
        .right = Picture1.ScaleWidth
        .top = 32
        .bottom = 46 'Picture1.ScaleHeight
    End With
    For Each st In s
        'insert first item at bottom left of view
        Picture1.Cls
        ' TextToPicture Picture1, Dir(st), eCenre
        'Picture1.Picture = Picture1.Image
        Picture1.Picture = FileExtractIcon(st)
        DrawText Picture1.hdc, Dir(st), -1, rct, 0 ' DT_WORDBREAK
        SavePicture Picture1.Image, App.Path & "\3DEngine\Meshes\icon" & CStr(z) & ".jpg"
        itemclass.Insert SetThing2d(0, 1, xcurrent, 110, zcurrent), "medkit.mesh", "icon" & CStr(z) & ".jpg", st
        z = z + 1
        'increment point to the right
        xcurrent = xcurrent + 5
        'if pointer is over bounds
        If xcurrent > xend Then
            'reset point to leftmost
            xcurrent = xstart
            'increment point upward
            zcurrent = zcurrent + 5
            'if pointer is above bound
            If zcurrent > zend Then
                'we are done this plane
                Exit For
            End If
        End If
    Next
    For Each st In s2
        'insert first item at bottom left of view
        Picture1.Cls
        Picture1.Picture = FileExtractIcon(st)
        DrawText Picture1.hdc, Dir(st), -1, rct, 0 ' DT_WORDBREAK
        SavePicture Picture1.Image, App.Path & "\3DEngine\Meshes\icon" & CStr(z) & ".jpg"
        itemclass.Insert SetThing2d(0, 1, xcurrent, 110, zcurrent), "medkit.mesh", "icon" & CStr(z) & ".jpg", st
        z = z + 1
        'increment point to the right
        xcurrent = xcurrent + 5
        'if pointer is over bounds
        If xcurrent > xend Then
            'reset point to leftmost
            xcurrent = xstart
            'increment point upward
            zcurrent = zcurrent + 5
            'if pointer is above bound
            If zcurrent > zend Then
                'we are done this plane
                Exit For
            End If
        End If
    Next
End Sub
Private Sub PrintText(StrText As String)
    On Error Resume Next
    Dim x As Long
    If Picture1.TextWidth(StrText) > Picture1.ScaleWidth Then
        x = InStr(1, StrReverse(StrText), " ")
        While Not x = 0
            If Picture1.TextWidth(left(StrText, Len(StrText) - x)) < Picture1.ScaleWidth Then
                StrText = left(StrText, Len(StrText) - x) & vbCrLf & right(StrText, x)
                x = 0
            Else
                x = InStr(x + 1, StrReverse(StrText), " ")
            End If
        Wend
    End If
    Picture1.Print StrText
End Sub
Friend Sub abort3Dxinput(ByVal thumbsdead As Boolean)
    On Error Resume Next
    xinputToDesktop = True
    RenderEnabled = False
    Dim ts As D3DVECTOR
    If thumbsdead = True Then
        RightThumbDead = True
        LeftThumbDead = True
    End If
    oldrv.x = 0
    oldrv.y = 0
    oldlv.x = 0
    oldlv.y = 0
    If xinpleft = True Then xinpleft = False
    If xinpright = True Then xinpright = False
    If xinpforward = True Then xinpforward = False
    If xinpbackward = True Then xinpbackward = False
    TargetSpeed = ts
End Sub
Private Function IsNoRedirectionBitmap(ByVal hWnd As Long) As Boolean
    On Error Resume Next
    If (apiGetWindowLong(hWnd, GWL_EXSTYLE) And WS_EX_NOREDIRECTIONBITMAP) = WS_EX_NOREDIRECTIONBITMAP Then IsNoRedirectionBitmap = True: Exit Function
    IsNoRedirectionBitmap = False
End Function
Private Function GetWinNameAPI(ByVal hWnd As Long, Optional ByVal gText As Boolean = True, Optional ByVal gClass As Boolean = True) As WINNAME
    On Error Resume Next
    Dim tLength As Long
    Dim rValue  As Long
    Dim rvalue2 As Long
    Dim n       As WINNAME
    n.lpText = "" ''''''''''''''''''''''''''''''''Initialize string for text name
    n.lpClass = "" '''''''''''''''''''''''''''''''Initialize string for class name
    If gText = True Then '''''''''''''''''''''''''If text is to be retrieved
        tLength = apiGetWindowTextLength(hWnd) + 4 'Get length
        n.lpText = Space(260) ' n.lpText.PadLeft(tLength) '''''Pad with buffer
        rValue = apiGetWindowText(hWnd, n.lpText, 260) 'Get text
        n.lpText = left(n.lpClass, rValue) '.lpText.Substring(0, rValue) 'Strip buffer
    End If
    If gClass = True Then ''''''''''''''''''''''''If class name is to be retrieved
        n.lpClass = Space(260) 'n.lpClass.PadLeft(260) '''''''Pad with buffer
        rvalue2 = apiGetClassName(hWnd, n.lpClass, 260) 'Get classname
        n.lpClass = left(n.lpClass, rvalue2) ' n.lpClass.Substring(0, rValue) 'Strip buffer
    End If
    GetWinNameAPI = n '''''''''''''''''''''''''''''''''''''Return WINNAME structure
End Function
Private Function GetWinName(ByVal hWnd As Long) As String
    On Error Resume Next
    Dim tLength As Long
    Dim rValue  As Long
    tLength = 260 'apiGetWindowTextLength(hWnd) + 4 'Get length
    GetWinName = Strings.Space(tLength) 'Pad with buffer
    rValue = apiGetWindowText(hWnd, GetWinName, tLength) 'Get text
    GetWinName = left(GetWinName, rValue) 'Strip buffer
End Function
Public Function SetStyleNoActivate(ByVal hWnd As Long) As Long
    On Error Resume Next: SetStyleNoActivate = 0
    Dim st As Long
    st = apiGetWindowLong(hWnd, GWL_EXSTYLE)
    If (st And WS_EX_NOACTIVATE) = WS_EX_NOACTIVATE Then SetStyleNoActivate = 0: Exit Function
    SetStyleNoActivate = apiSetWindowLong(hWnd, GWL_EXSTYLE, st Or WS_EX_NOACTIVATE)
End Function
Private Sub UNLOCK_XINPUT()
    On Error Resume Next
    xenable = False
    IsXINPUTSupported = False
    Dim hWnd As Long
    hWnd = apiGetForegroundWindow
    If hWnd = Me.hWnd Then Exit Sub
    If IsNoRedirectionBitmap(hWnd) = True Then SetStyleNoActivate (hWnd)
    Dim twnd As Long
    twnd = apiFindWindow("Shell_TrayWnd", vbNullString)
    If apiIsWindow(twnd) <> 0 Then apiSetForegroundWindow (twnd)
    Dim pwnd As Long
    pwnd = apiFindWindow("Progman", "Program Manager")
    If apiIsWindow(pwnd) = 0 Then pwnd = apiSetForegroundWindow(apiGetDesktopWindow)
    If apiIsWindow(pwnd) <> 0 Then apiSetForegroundWindow (pwnd)
End Sub
Private Sub Isxinput(Optional ByVal hWnd As Long = 0)
    On Error Resume Next
    If apiIsWindow(hWnd) = 0 Then hWnd = apiGetForegroundWindow
    If IsNoRedirectionBitmap(hWnd) = True Then
        Dim wn As WINNAME
        wn = GetWinNameAPI(hWnd, False, True)
        wn.lpText = GetWinName(hWnd)
        If wn.lpClass = "ApplicationFrameWindow" Then
            IsXINPUTSupported = False
            If LCase(wn.lpText) = "xbox" Then
                xenable = True
                modifierkeydown = False
            Else
                xenable = False
            End If
            SetStyleNoActivate (hWnd) 'If it appears to be a metro store app window with the rendering style of a store app (or taskbar child core window) then'redirect input to mouse only since this rendering style would conflict with internal "focus navigation"
            UNLOCK_XINPUT
        ElseIf wn.lpClass = "Windows.UI.Core.CoreWindow" Then
            xenable = False
            If wn.lpText = "Volume Control" Then
                IsXINPUTSupported = False
            ElseIf wn.lpText = "Start" Then
                IsXINPUTSupported = True
            ElseIf wn.lpText = "Cortana" Then
                IsXINPUTSupported = True
            ElseIf hWnd = wn.lpText = "Action Center" Then
                IsXINPUTSupported = True
            ElseIf hWnd = wn.lpText = "Windows Shell Experience Host" Then
                IsXINPUTSupported = True
            Else
                IsXINPUTSupported = True
            End If
        Else
            xenable = False
            IsXINPUTSupported = False
        End If
    Else
        xenable = False
        IsXINPUTSupported = False
    End If
End Sub
Private Function GetStringValue(ByRef searchString As String, ByVal lVerPointer As Long, ByVal lBufferLen As Long, ByRef sBuffer() As Byte) As String
    On Error Resume Next
    Dim Buffer  As String
    Dim strTemp As String
    Dim RC      As Long
    GetStringValue = ""
    Buffer = String(260, 0)
    RC = apiVerQueryValueByteLong(sBuffer(0), searchString, lVerPointer, lBufferLen)
    If RC <> 0 Then
        Call apilstrcpy(Buffer, lVerPointer)
        GetStringValue = Mid$(Buffer, 1, InStr(Buffer, Chr(0)) - 1)
    End If
End Function
Private Function GetTarget(ByVal FileName As String) As String
    On Error Resume Next
    Dim obj As Object, Shortcut As Object
    Set obj = CreateObject("WScript.Shell")
    Set Shortcut = obj.CreateShortcut(FileName)
    GetTarget = LCase(Shortcut.TargetPath)
    Set obj = Nothing
    Set Shortcut = Nothing
End Function
Private Function DoesContain(ByVal Inp As String) As Boolean
    Dim v As Long
    '    For v = 1 To ListView.ListCount
    '        If ListView1.ListItems.Item(v).SubItems(1) = Inp Then
    '            DoesContain = True
    '            Exit Function
    '        End If
    '    Next
End Function
Public Function RunShellExecute(ByVal sTopic As String, ByVal sFile As String, ByVal sParams As String, ByVal sDirectory As String, ByVal nShowCmd As Long) As Long
    On Error Resume Next
    RunShellExecute = apiShellExecute(0, sTopic, sFile, sParams, sDirectory, nShowCmd)
End Function
Private Function CheckVersion() As Long
    Dim tOS As OSVERSIONINFO
    tOS.dwOSVersionInfoSize = Len(tOS)
    Call apiGetVersionEx(tOS)
    CheckVersion = tOS.dwPlatformID
End Function
Private Function GetEXEProcessID(ByVal sEXE As String) As Long
    On Error Resume Next
    Dim aPID()     As Long
    Dim lProcesses As Long
    Dim lprocess   As Long
    Dim lModule    As Long
    Dim sName      As String
    Dim iIndex     As Integer
    Dim bCopied    As Long
    Dim lSnapShot  As Long
    Dim tPE        As PROCESSENTRY32
    Dim bDone      As Boolean
    If CheckVersion() = VER_PLATFORM_WIN32_WINDOWS Then
        lSnapShot = apiCreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
        If lSnapShot < 0 Then Exit Function
        tPE.dwSize = Len(tPE)
        bCopied = apiProcess32First(lSnapShot, tPE)
        Do While bCopied
            sName = left(tPE.szExeFile, InStr(tPE.szExeFile, Chr(0)) - 1)
            sName = Mid(sName, InStrRev(sName, "\") + 1)
            If InStr(sName, Chr(0)) Then
                sName = left(sName, InStr(sName, Chr(0)) - 1)
            End If
            bCopied = apiProcess32Next(lSnapShot, tPE)
            If StrComp(sEXE, sName, vbTextCompare) = 0 Then
                GetEXEProcessID = tPE.th32ProcessID
                Exit Do
            End If
        Loop
    Else
        ReDim aPID(255)
        Call apiEnumProcesses(aPID(0), 1024, lProcesses)
        lProcesses = lProcesses / 4
        ReDim Preserve aPID(lProcesses)
        For iIndex = 0 To lProcesses - 1
            lprocess = apiOpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, aPID(iIndex))
            If lprocess <> 0 Then
                If apiEnumProcessModules(lprocess, lModule, 4, 0&) Then
                    sName = Space(260)
                    Call apiGetModuleFileNameExA(lprocess, lModule, sName, Len(sName))
                    If InStr(sName, "\") > 0 Then sName = Mid(sName, InStrRev(sName, "\") + 1)
                    If InStr(sName, Chr(0)) Then sName = left(sName, InStr(sName, Chr(0)) - 1)
                    If StrComp(sEXE, sName, vbTextCompare) = 0 Then
                        GetEXEProcessID = aPID(iIndex)
                        bDone = True
                    End If
                End If
                apiCloseHandle lprocess
                If bDone Then Exit For
            End If
        Next
    End If
End Function
Friend Function TerminateEXE(ByVal sEXE As String) As Boolean
    On Error Resume Next
    Dim pid As Long
    Dim lprocess As Long
    Dim wts As New clsProcessWTS
    pid = wts.GetWTSPID(sEXE)
    If pid = 0 Then Exit Function
    lprocess = apiOpenProcess(PROCESS_TERMINATE, 0, pid)
    If lprocess = 0 Then Exit Function
    apiTerminateProcess lprocess, 0
    apiCloseHandle lprocess
    TerminateEXE = True
End Function
Private Sub CenterCursor()
    On Error Resume Next
    BBWidth = My.Computer.Screen.PrimaryScreen.Bounds.Width_
    BBHeight = My.Computer.Screen.PrimaryScreen.Bounds.Height
    ScrCenterX = BBWidth \ 2
    ScrCenterY = BBHeight \ 2
End Sub
Private Function IsAppWindowForeground() As Boolean
    On Error Resume Next
    Dim b As Boolean
    b = False
    Dim hWnd As Long
    hWnd = apiGetForegroundWindow
    If apiIsWindow(hWnd) <> 0 Then
        If hWnd = apiFindWindow("Windows.UI.Core.CoreWindow", vbNullString) Then b = True
        If hWnd = apiFindWindow("ApplicationFrameWindow", vbNullString) Then b = True
    End If
    IsAppWindowForeground = b
End Function
Private Function IsAppFrameWindowForeground() As Boolean
    On Error Resume Next
    Dim b As Boolean
    b = False
    Dim hWnd As Long
    hWnd = apiGetForegroundWindow
    If apiIsWindow(hWnd) <> 0 Then
        If hWnd = apiFindWindow("ApplicationFrameWindow", vbNullString) Then b = True
    End If
    IsAppFrameWindowForeground = b
End Function
Private Sub resizetoPrimaryscreen()
    On Error Resume Next
    Dim ds As New clsDriveSerial
    Dim dn As String
    Dim d As Long
    dn = ds.GetDesktopName
    If LCase(dn) = "default" Then
    Else
        d = 8
    End If
    Me.left = ScaleX(My.Computer.Screen.PrimaryScreen.WorkingArea.left, vbPixels, vbTwips)
    Me.top = ScaleX(My.Computer.Screen.PrimaryScreen.WorkingArea.top, vbPixels, vbTwips)
    Me.Width = ScaleX(My.Computer.Screen.PrimaryScreen.WorkingArea.Width_, vbPixels, vbTwips)
    Me.Height = ScaleY(My.Computer.Screen.PrimaryScreen.WorkingArea.Height + d, vbPixels, vbTwips)
End Sub
Public Function SetWindowPos(ByVal hWnd As Long, ByVal insAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, Optional ByVal show As Boolean = False, Optional ByVal Activate As Boolean = False) As Long
    On Error Resume Next
    Dim flgs As Long
    flgs = SWP_NOMOVE + SWP_NOSIZE
    If show = True Then flgs = flgs + SWP_SHOWWINDOW
    If Activate = False Then flgs = flgs + SWP_NOACTIVATE
    flgs = flgs + SWP_NOOWNERZORDER
    SetWindowPos = apiSetWindowPos(hWnd, insAfter, x, y, cx, cy, flgs)
End Function
Private Function SetThing2d(ByVal tg As Byte, ByVal typ As Byte, ByVal x As Integer, ByVal z As Integer, ByVal y As Integer) As tThing2D
    On Error Resume Next
    Dim t As tThing2D ' 0 1 234 255 28
    t.Tag = tg
    t.tType = typ
    t.x = x
    t.y = y
    t.z = z
    SetThing2d = t
End Function
Private Sub SelectDeskIcon()
    On Error Resume Next
    Dim hWnd As Long
    Dim cwnd As Long
    Dim fwnd As Long
    hWnd = apiFindWindow("Progman", "Program Manager")
    If hWnd = 0 Then Exit Sub
    cwnd = apiFindWindowEx(hWnd, 0, "SHELLDLL_DefView", "")
    If cwnd = 0 Then Exit Sub
    fwnd = apiFindWindowEx(cwnd, 0, "SysListView32", "FolderView")
End Sub
Private Function WindowFromPoint(ByRef p As POINTAPI) As Long
    On Error Resume Next
    WindowFromPoint = apiWindowFromPoint(p.x, p.y)
End Function
Private Function GetForegroundFromPoint(ByRef p As POINTAPI) As Long
    On Error Resume Next
    Dim wfp As Long
    wfp = WindowFromPoint(p)
    If wfp = 0 Then Exit Function
    GetForegroundFromPoint = apiGetAncestor(wfp, GA_ROOT)
End Function
Private Sub pooper()
    On Error Resume Next
    Const ID_ABOUT     As Long = 101
    Const ID_SEPARATOR As Long = 102
    Const ID_EXIT      As Long = 103
    Dim hPopupMenu     As Long    ' handle to the popup menu to display
    Dim mii            As MENUITEMINFO   ' describes menu items to add
    Dim curpos         As POINT_TYPE  ' holds the current mouse coordinates
    Dim menusel        As Long       ' ID of what the user selected in the popup menu
    Dim RetVal         As Long        ' generic return value
    ' Create the popup menu that will be displayed.
    hPopupMenu = CreatePopupMenu()
    ' Add the menu's first item: "About This Problem..."
    With mii
        ' The size of this structure.
        .cbSize = Len(mii)
        ' Which elements of the structure to use.
        .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
        ' The type of item: a string.
        .fType = MFT_STRING
        ' This item is currently enabled and is the default item.
        .fState = MFS_ENABLED Or MFS_DEFAULT
        ' Assign this item an item identifier.
        .wID = ID_ABOUT
        ' Display the following text for the item.
        .dwTypeData = "&About This Example..."
        .cch = Len(.dwTypeData)
    End With
    RetVal = InsertMenuItem(hPopupMenu, 0, 1, mii)
    ' Add the second item: a separator bar.
    With mii
        .fType = MFT_SEPARATOR
        .fState = MFS_ENABLED
        .wID = ID_SEPARATOR
    End With
    RetVal = InsertMenuItem(hPopupMenu, 1, 1, mii)
    ' Add the final item: "Exit".
    With mii
        .fType = MFT_STRING
        .wID = ID_EXIT
        .dwTypeData = "E&xit"
        .cch = Len(.dwTypeData)
    End With
    RetVal = InsertMenuItem(hPopupMenu, 2, 1, mii)
    ' Determine where the mouse cursor currently is, in order to have
    ' the popup menu appear at that point.
    RetVal = GetCursorPos(curpos)
    ' Display the popup menu at the mouse cursor.  Instead of sending messages
    ' to window Form1, have the function merely return the ID of the user's selection.
    menusel = TrackPopupMenu(hPopupMenu, TPM_TOPALIGN Or TPM_LEFTALIGN Or TPM_NONOTIFY Or TPM_RETURNCMD Or TPM_LEFTBUTTON, curpos.x, curpos.y, 0, Me.hWnd, 0)
    ' Before acting upon the user's selection, destroy the popup menu now.
    RetVal = DestroyMenu(hPopupMenu)
    Select Case menusel
        Case ID_ABOUT
            ' Use the Visual Basic MsgBox function to display a short message in a dialog
            ' box.  Using the MessageBox API function isn't necessary.
            RetVal = MsgBox("This example demonstrates how to use the API to display " & "a pop-up menu.", vbOKOnly Or vbInformation, "Windows API Guide")
        Case ID_EXIT
            ' End this program by closing and unloading Form1.
            ' Unload Form1
    End Select
End Sub
Private Function IsSizable(ByVal hWnd As Long) As Boolean
    On Error Resume Next
    Dim st As Long
    st = apiGetWindowLong(hWnd, GWL_STYLE)
    If (st And WS_SIZEBOX) = WS_SIZEBOX Then
        IsSizable = True
    End If
End Function
Private Function IsMaximizable(ByVal hWnd As Long) As Boolean
    On Error Resume Next
    Dim st As Long
    st = apiGetWindowLong(hWnd, GWL_STYLE)
    If (st And WS_MAXIMIZEBOX) = WS_MAXIMIZEBOX Then
        IsMaximizable = True
    End If
End Function '
Private Function IsMinimizable(ByVal hWnd As Long) As Boolean
    On Error Resume Next
    Dim st As Long
    st = apiGetWindowLong(hWnd, GWL_STYLE)
    If (st And WS_MINIMIZEBOX) = WS_MINIMIZEBOX Then
        IsMinimizable = True
    End If
End Function
Private Function VirtualScreenWidth() As Long
    On Error Resume Next
    VirtualScreenWidth = apiGetSystemMetrics(SM_CXVIRTUALSCREEN)
End Function
Private Function VirtualScreenHeight() As Long
    On Error Resume Next
    VirtualScreenHeight = apiGetSystemMetrics(SM_CYVIRTUALSCREEN)
End Function
Private Function DisplayMonitorCount() As Long
    On Error Resume Next
    DisplayMonitorCount = apiGetSystemMetrics(SM_CMONITORS)
End Function
Private Function AllMonitorsSame() As Long
    On Error Resume Next
    AllMonitorsSame = apiGetSystemMetrics(SM_SAMEDISPLAYFORMAT)
End Function
Private Function AppDoEvents() As Boolean
    On Error Resume Next
    AppDoEvents = False
    Dim Message As Msg
    If apiGetMessage(Message, 0, 0, 0) <> 0 Then ' use the GetMessage version if you only want to do processing if there's a message
        Dim ret  As Long
        Dim ret2 As Long
        ret = apiTranslateMessage(Message)
        ret2 = apiDispatchMessage(Message)
        If ret <> 0 And ret2 <> 0 Then AppDoEvents = True 'signal that there was a message and it was processes successfully
    End If
End Function
Private Sub SaveGame(s As String)
    On Error Resume Next
    Dim nf As Long
    nf = FreeFile
    Open s For Binary As #nf
    itemclass.Save nf
    PhysSave nf
    buttonClass.Save nf
    Close #nf
End Sub
Private Sub LoadGame(s As String)
    On Error Resume Next
    Dim nf As Long
    nf = FreeFile
    Open s For Binary As #nf
    itemclass.Load nf
    PhysLoad nf
    buttonClass.Load nf
    '    targetClass.Reset
    soundClass.StopPlay
    Close #nf
End Sub
'Private Sub SaveSettings()
'    On Error Resume Next
'    Dim nf As Integer
'    Dim v  As Long
'    nf = FreeFile
'    Open App.Path & "\hotkey.txt" For Binary As #nf
'    Put #nf, , KeyJump
'    Put #nf, , KeyCrouch
'    Put #nf, , KeyFire
'    Put #nf, , KeyUse
'    Put #nf, , KeyScreenshot
'    Close #nf
'End Sub
'Private Sub LoadSettings()
'    On Error Resume Next
'    Dim nf As Integer
'    Dim v  As Long
'    nf = FreeFile
'    Open App.Path & "\hotkey.txt" For Binary As #nf
'    Get #nf, , KeyJump
'    Get #nf, , KeyCrouch
'    Get #nf, , KeyFire
'    Get #nf, , KeyUse
'    Get #nf, , KeyScreenshot
'    Close #nf
'End Sub
Private Function IsFree() As String
    On Error Resume Next
    Dim i As Long
    Dim s As String
    Do
        i = i + 1
        s = ""
        s = App.Path & "\screenshots\" & CStr(i) & ".jpg"
        If Dir$(s) = "" Then IsFree = s: Exit Do
    Loop
End Function
Private Sub SetCamera()
    On Error Resume Next
    Dim v As D3DVECTOR
    PlDir = Vec3(SinA * CosD, -SinD, CosA * CosD)
    PlUp = Vec3(SinA * SinD, CosD, CosA * SinD)
    PlRight = Vec3(CosA, 0, -SinA)
    Vec3Add v, PlPos, PlDir
    MatrixLookAtLH mView, PlPos, v, Vec3(0, 1, 0)
End Sub
Private Sub QTimeReset(ByVal time As Double)
    On Error Resume Next
    MaxCur = CCur("922337203685477") + CCur(0.5807)
    MinCur = -MaxCur - CCur(0.0001)
    apiQueryPerformanceCounter OldQC
    apiQueryPerformanceFrequency QF
    QTimeVal = time
End Sub
Private Sub MyDoEvents()
    On Error Resume Next
    Dim Msg As Msg
    Do
        If apiPeekMessage(Msg, 0, 0, 0, PM_REMOVE) = 0 Then Exit Do
        apiTranslateMessage Msg
        apiDispatchMessage Msg
    Loop
End Sub
Private Function QTime() As Double
    On Error Resume Next
    Dim QC As Currency
    apiQueryPerformanceCounter QC
    If QC >= OldQC Then
        QTimeVal = QTimeVal + (QC - OldQC) / QF
    Else
        QTimeVal = QTimeVal + ((MaxCur - OldQC) + (QC - MinCur)) / QF
    End If
    OldQC = QC
    QTime = QTimeVal
End Function
Private Sub PhysInit()
    On Error Resume Next
    Dim i As Long
    PlPos = StartPos
    PlAngle = 0
    PlDiff = 0
    PlSpeed = Vec3(0, 0, 0)
    TargetSpeed = Vec3(0, 0, 0)
    directxGravity = -0.001
    UnCrouch
End Sub
Private Sub PhysSave(ByVal nf As Long)
    On Error Resume Next
    Put #nf, , PlPos
    Put #nf, , PlSpeed
    Put #nf, , PlDir
    Put #nf, , PlUp
    Put #nf, , PlRight
    Put #nf, , PlAngle
    Put #nf, , PlDiff
End Sub
Private Sub PhysLoad(ByVal nf As Long)
    On Error Resume Next
    Get #nf, , PlPos
    Get #nf, , PlSpeed
    Get #nf, , PlDir
    Get #nf, , PlUp
    Get #nf, , PlRight
    Get #nf, , PlAngle
    Get #nf, , PlDiff
End Sub
Private Sub SetTargetSpeed(ts As D3DVECTOR, xy As Double)
    On Error Resume Next
    Dim t As Single
    t = Vec3Length(ts)
    If PlIsFly Then
        If t > 0.1 Then
            Vec3Scale TargetSpeed, ts, MaxSpeedFly / t
        Else
            TargetSpeed.y = 0
            TargetSpeed.x = PlSpeed.x
            TargetSpeed.z = PlSpeed.z
        End If
    Else
        If t > 0.1 Then
            Dim RT As Double
            RT = oldis.gamepad.bLeftTrigger
            If RT > 0 Then RT = (RT / 255)
            RT = RT * pointerMaxWalkAcceleration
            Vec3Scale TargetSpeed, ts, (MaxSpeed / t) * (pointerMaxWalkSpeed * (1 + RT)) * (xy)
        Else
            TargetSpeed = Vec3(0, 0, 0)
        End If
    End If
End Sub
Private Sub UnCrouch()
    PlHeight = UnCrouchHeight
    PlIsCrouch = False
End Sub
Private Sub Crouch()
    PlHeight = CrouchHeight
    PlIsCrouch = True
End Sub
Friend Sub Respawn()
    On Error Resume Next
    RenderEnabled = False
    xinputToDesktop = True
    PhysInit
    TerminateEXE "WindowContextMenu.exe"
    LoadGame App.Path & "\Game0.sav"
    LoadGame App.Path & "\Game1.sav"
    oldfls = ""
End Sub
Private Sub PhysTick()
    On Error Resume Next
    Dim i          As Long
    Dim v          As D3DVECTOR
    Dim vP         As D3DVECTOR
    Dim vN         As D3DVECTOR
    Dim s          As Single
    Dim t          As Long
    Dim f          As Boolean
    Dim Frc        As Single
    Dim Floor      As Single
    Dim fl         As Single
    Static OldPos  As D3DVECTOR
    Static vNFloor As D3DVECTOR
    Static Stuck   As Long
    If PlPos.y < -16 Then
        soundClass.Damage
        Respawn
        Exit Sub
    End If
    v = TargetSpeed
    s = Sqr(vNFloor.x * vNFloor.x + vNFloor.z * vNFloor.z)
    If s > 0.7 Then
        Vec3Cross vN, vNFloor, Vec3(0, 1, 0)
        Vec3Cross vN, vN, vNFloor
        Vec3Scale vN, vN, -2 * (s - 0.7)
        Vec3Add v, v, vN
        v.y = 0
        WantJump = False
    End If
    If (Not PlIsFly) And WantJump And PlSpeed.y > -0.1 Then
        PlSpeed.y = JumpSpeedY
        PlSpeed.x = PlSpeed.x + TargetSpeed.x * Abs(PlSpeed.x)
        PlSpeed.z = PlSpeed.z + TargetSpeed.z * Abs(PlSpeed.z)
        WantJump = False
        soundClass.Step 0.5
    End If
    v.x = v.x - PlSpeed.x
    v.z = v.z - PlSpeed.z
    s = Sqr(v.x * v.x + v.z * v.z)
    If PlIsFly Then
        Frc = MaxForceFly
    ElseIf PlIsCrouch Then
        Frc = MaxForceCrouch
    Else
        Frc = MaxForceUnCrouch
    End If
    If s > Frc Then
        Vec3Scale PlForce, v, Frc / s
    Else
        PlForce = v
    End If
    Vec3Add PlForce, PlForce, PlFlFrc
    PlForce.y = PlForce.y + directxGravity
    Vec3Add PlSpeed, PlSpeed, PlForce
    Floor = -999999
    vNFloor = Vec3(0, 1, 0)
    For t = 0 To 3
        If t = 2 Then Vec3Subtract PlSpeed, PlSpeed, PlForce
        f = True
        Vec3Add vP, PlPos, PlSpeed
        Vec3Add vP, vP, AddPos
        LS.GetHeight vP.x, vP.z, fl
        If Floor < fl Then
            Floor = fl
            LS.GetNorm vP.x, vP.z, vNFloor
        End If
        If vP.y < Floor + CrouchHeight - 0.5 Then
            f = False
            s = Vec3Dot(PlSpeed, vNFloor)
            If s < 0 Then
                Vec3Scale v, vNFloor, -1.001 * s
            Else
                Vec3Scale v, vNFloor, 0.001 * s
            End If
            Vec3Add PlSpeed, PlSpeed, v
            Vec3Add vP, PlPos, PlSpeed
            Vec3Add vP, vP, AddPos
        Else
            For i = 0 To SectorCnt - 1
                If Sector(i).PlayerInSector(vP, PlPos, 1.125, CrouchHeight - 0.5, HeadHeight, PlHeight, vN, fl) = True Then
                    f = False
                    s = Vec3Dot(PlSpeed, vN)
                    If s < 0 Then
                        Vec3Scale v, vN, -1.001 * s
                    Else
                        Vec3Scale v, vN, 0.001 * s
                    End If
                    Vec3Add PlSpeed, PlSpeed, v
                    Vec3Add vP, PlPos, PlSpeed
                    Vec3Add vP, vP, AddPos
                    If vN.y < 0 Then
                        If PlHeight > CrouchHeight Then PlHeight = PlHeight - 0.01
                        PlPos.y = PlPos.y - 0.03
                    End If
                End If
                If Floor < fl Then
                    Floor = fl
                    vNFloor = vN
                End If
            Next i
        End If
        AddPos.x = 0
        AddPos.z = 0
        If f Then
            PlPos = vP
            If PlPos.y - Floor < PlHeight + 0.5 Then
                s = Floor + PlHeight - PlPos.y
                PlFlFrc.y = s * 0.01 - PlSpeed.y * 0.27
                If PlFlFrc.y > 0.007 Then
                    PlFlFrc.y = 0.007
                ElseIf PlFlFrc.y < 0 Then
                    PlFlFrc.y = 0
                End If
                If PlIsFly Then
                    s = Abs(PlSpeed.y) * 5
                    If s > 1 Then s = 1
                    soundClass.Step s
                End If
                PlIsFly = False
                MaxSpeed = MaxSpeedCrouch + (MaxSpeedUnCrouch - MaxSpeedCrouch) * (PlPos.y - CrouchHeight - Floor) / (UnCrouchHeight - CrouchHeight)
            Else
                PlFlFrc.y = 0
                PlIsFly = True
                vNFloor = Vec3(0, 0, 0)
            End If
            Stuck = 0
            OldPos = PlPos
            Exit Sub
        End If
    Next t
    If Stuck <= 100 Then
        Stuck = Stuck + 1
        Vec3Scale PlSpeed, PlSpeed, -0.000005
        PlPos = OldPos
    End If
End Sub
Friend Sub D3DInit(ByVal hWnd As Long)
    On Error Resume Next
    Set D3D = CreateDirect3D
    d3dpp.SwapEffect = D3DSWAPEFFECT_DISCARD
    d3dpp.BackBufferCount = 1
    If directxVSync = True Then
        d3dpp.PresentationInterval = D3DPRESENT_ONE
    Else
        d3dpp.PresentationInterval = D3DPRESENT_IMMEDIATE
    End If
    d3dpp.EnableAutoDepthStencil = D3D_TRUE
    d3dpp.AutoDepthStencilFormat = D3DFMT_D24S8
    d3dpp.Windowed = D3D_TRUE
    d3dpp.BackBufferWidth = BBWidth
    d3dpp.BackBufferHeight = BBHeight
    d3dpp.BackBufferFormat = D3DFMT_A8R8G8B8
    ' directAspect = BBWidth / BBHeight
'    directxAspect = (Screen.Width \ Screen.TwipsPerPixelX) / (Screen.Height \ Screen.TwipsPerPixelY)
    Set Dev = D3D.CreateDevice(hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING Or D3DCREATE_FPU_PRESERVE, d3dpp)
    If Dev Is Nothing Then
        Set Dev = D3D.CreateDevice(hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING Or D3DCREATE_FPU_PRESERVE, d3dpp)
    End If
    MatrixPerspectiveFovLH mProj, directxFovY, directxAspect, 0.1, 400
    ZOptimize = True
    ResetStates
End Sub
Friend Sub ResetStates()
    On Error Resume Next
    If RenderEnabled = True Then
        frmMain.shadowClass.KillSurf
        Dev.Reset d3dpp
        frmMain.shadowClass.CreateSurf
    End If
    Dev.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE
    Dev.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    Dev.SetRenderState D3DRS_LIGHTING, D3D_FALSE
    TexFilter 0, directxTexFilters, directxAnisotropy
    TexFilter 1, directxTexFilters, directxAnisotropy
    TexFilter 2, directxTexFilters, directxAnisotropy
    TexFilter 3, TextureFilter_BiLinear
End Sub
Friend Sub TexFilter(ByVal Stage As Long, ByVal TF As TextureFilter, Optional ByVal MaxAnisotropy As Long = 2)
    On Error Resume Next
    Select Case TF
        Case TextureFilter_BiLinear
            Dev.SetSamplerState Stage, D3DSAMP_MIPFILTER, D3DTEXF_POINT
            Dev.SetSamplerState Stage, D3DSAMP_MAGFILTER, D3DTEXF_LINEAR
            Dev.SetSamplerState Stage, D3DSAMP_MINFILTER, D3DTEXF_LINEAR
        Case TextureFilter_TriLinear
            Dev.SetSamplerState Stage, D3DSAMP_MIPFILTER, D3DTEXF_LINEAR
            Dev.SetSamplerState Stage, D3DSAMP_MAGFILTER, D3DTEXF_LINEAR
            Dev.SetSamplerState Stage, D3DSAMP_MINFILTER, D3DTEXF_LINEAR
        Case TextureFilter_Anisotropic
            Dev.SetSamplerState Stage, D3DSAMP_MIPFILTER, D3DTEXF_LINEAR
            Dev.SetSamplerState Stage, D3DSAMP_MAGFILTER, D3DTEXF_LINEAR
            Dev.SetSamplerState Stage, D3DSAMP_MINFILTER, D3DTEXF_ANISOTROPIC
            Dev.SetSamplerState Stage, D3DSAMP_MAXANISOTROPY, MaxAnisotropy
    End Select
End Sub
Friend Sub D3DTerminate()
    On Error Resume Next
    Set Dev = Nothing
    Set D3D = Nothing
End Sub
'Private Sub fnFolderItemVerbsVB()
'    '    Dim objShell   As Shell
'    '    Dim objFolder2 As Folder2
'    '    Dim ssfWINDOWS As Long
'    '
'    '    ssfWINDOWS = 36
'    '    Set objShell = New Shell
'    '    Set objFolder2 = objShell.NameSpace(ssfWINDOWS)
'    '        If (Not objFolder2 Is Nothing) Then
'    '            Dim objFolderItem As FolderItem
'    '
'    '            Set objFolderItem = objFolder2.Self
'    '                If (Not objFolderItem Is Nothing) Then
'    '                    Dim objItemVerbs As FolderItemVerbs
'    '
'    '                    Set objItemVerbs = objFolderItem.Verbs
'    '                        If (Not objItemVerbs Is Nothing) Then
'    '                            'Add code here
'    '                        End If
'    '                    Set objItemVerbs = Nothing
'    '                Else
'    '                    'FolderItem object returned nothing.
'    '                End If
'    '            Set objFolderItem = Nothing
'    '        Else
'    '            'Folder object returned nothing.
'    '        End If
'    '    Set objFolder2 = Nothing
'    '    Set objShell = Nothing
'End Sub
'Friend Sub LoadPinnedStartMenu()
'    On Error Resume Next
'    Dim pth As String
'    Dim i   As Long
'    Dim sp  As StdPicture
'    Dim k   As Long
'   ' ListView2.ListItems.Clear
'    ImageList2.ListImages.Clear
'    Dim finf As String
'    For i = 1 To 100
'        pth = GetSetting("WindowLauncher", "PinnedStartMenu", "Path" & CStr(i), "")
'        If pth = "" Or Dir(pth, vbNormal) = "" Then
'        Else
'            finf = frmStartMenu.GetFileInformation(pth)
'            If finf = "" Then
'                finf = Dir(f, vbNormal)
'            End If
'            k = k + 1
'            Set sp = Nothing
'            Set sp = FileExtractIcon(pth)
'            If sp Is Nothing Then
'                ImageList2.ListImages.Add , , picRunDialog.Picture
'            Else
'                ImageList2.ListImages.Add , , sp
'               ' ListView2.ListItems.Add(, , finf, , k).SubItems(1) = pth
'            End If
'        End If
'    Next
'End Sub
'Private Sub drawicontopicDC(ByVal i As Long)
'
'    Dim hIML As IUnknown
'    Dim GUID(0 To 3) As Long
'    Dim lResult As Long
'
'    Dim lIconSize As Long
'    lIconSize = SHIL_JUMBO
'
'    If IIDFromString(StrPtr(IID_IImageList), GUID(0)) = 0 Then
'        On Error Resume Next
'        lResult = SHGetImageList(lIconSize, GUID(0), ByVal VarPtr(hIML))
''        Select Case lResult
''        Case 0&
''            If Err Then
''                Err.Clear
''                lResult = SHGetImageListXP(lIconSize, GUID(0), ByVal VarPtr(hIML))
''                If Err Then lResult = E_INVALIDARG ' assign any non-zero value; function not exported
''            End If
''        Case E_INVALIDARG
''            ' possibly calling API with SHIL_JUMBO on XP?
''        Case Else
''            ' some other error
''        End Select
''        On Error GoTo 0
'        If lResult = 0& Then
'
'            ' assume you have the icon index you want; here we'll just use 5
'            Dim hIcon As Long
'            hIcon = ImageList_GetIcon(ObjPtr(hIML), i, 0)
'            If hIcon Then
'                DrawIconEx Picture1.hDC, 0, 0, hIcon, 0, 0, 0, 0, 3
'
'
'                DestroyIcon hIcon
'            End If
'        End If
'    End If
'
'End Sub
