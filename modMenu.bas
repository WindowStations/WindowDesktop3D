Attribute VB_Name = "modMenu"
Option Explicit
Const TPM_CENTERALIGN       As Long = &H4
Const TPM_RIGHTALIGN        As Long = &H8
Const TPM_BOTTOMALIGN       As Long = &H20
Const TPM_VCENTERALIGN      As Long = &H10
Const TPM_RIGHTBUTTON       As Long = &H2
'Const TPM_LEFTALIGN  As Long = &H0
'Const TPM_TOPALIGN  As Long = &H0
'Const TPM_NONOTIFY  As Long = &H80
'Const TPM_RETURNCMD  As Long = &H100
'Const TPM_LEFTBUTTON  As Long = &H0
Public Const TPM_LEFTALIGN  As Long = &H0
Public Const TPM_TOPALIGN   As Long = &H0
Public Const TPM_NONOTIFY   As Long = &H80
Public Const TPM_RETURNCMD  As Long = &H100
Public Const TPM_LEFTBUTTON As Long = &H0
Public Const MIIM_STATE     As Long = &H1
Public Const MIIM_ID        As Long = &H2
Public Const MIIM_TYPE      As Long = &H10
Public Const MFT_SEPARATOR  As Long = &H800
Public Const MFT_STRING     As Long = &H0
Public Const MFS_DEFAULT    As Long = &H1000
Public Const MFS_ENABLED    As Long = &H0
Public Type MENUITEMINFO
   cbSize As Long
   fMask As Long
   fType As Long
   fState As Long
   wID As Long
   hSubMenu As Long
   hbmpChecked As Long
   hbmpUnchecked As Long
   dwItemData As Long
   dwTypeData As String
   cch As Long
End Type
Public Type POINT_TYPE
   x As Long
   y As Long
End Type
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, ByRef lpmii As MENUITEMINFO) As Long
Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal uFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As Long, ByVal prcRect As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINT_TYPE) As Long
