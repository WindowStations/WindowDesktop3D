Attribute VB_Name = "modEventHook"
'Private Const WINEVENT_SKIPOWNPROCESS As Long = 2
'Private Const WINEVENT_SKIPOWNTHREAD As Long = 1
'Private Const WINEVENT_INCONTEXT As Long = 4
'Private Const WINEVENT_32BITCALLER As Long = 32768
'Private Const WINEVENT_VALID As Long = 32775
'Private Const SYS_SOUND As Long = 1
'Private Const SYS_ALERT As Long = 2
'Private Const EVENT_OB_FOCUS As Long = 32773
'Private Const EVENT_OB_SELECTION As Long = 32774
'Private Const SYS_FOREGROUND As Long = 3 'active window changing.
'Private Const SYS_MENUSTART As Long = 4 '--entering menu mode.
'Private Const SYS_MENUEND As Long = 5 '--leaving menu mode.
'Private Const SYS_MENUPOPUPSTART As Long = 6 '--popup menu popping up.
'Private Const SYS_MENUPOPUPEND As Long = 7 'popup menu about to close.
'Private Const SYS_CAPTURESTART As Long = 8 '--window took capture.
'Private Const SYS_CAPTUREEND As Long = 9 '--window released capture.
'Private Const SYS_MOVESIZESTART As Long = &HA '-- beginning move or resize.
'Private Const SYS_MOVESIZEEND As Long = &HB '--ending move or resize.
'Private Const SYS_CONTEXTHELPSTART As Long = &HC
'Private Const SYS_CONTEXTHELPEND As Long = &HD
'Private Const SYS_DRAGDROPSTART As Long = &HE
'Private Const SYS_DRAGDROPEND As Long = &HF
'Private Const SYS_DIALOGSTART As Long = &H10 '--dialogue opens.
'Private Const SYS_DIALOGEND As Long = &H11 '--dialogue about to close.
'Private Const SYS_SCROLLINGSTART As Long = &H12 '--scrolling.
'Private Const SYS_SCROLLINGEND As Long = &H13
'Private Const SYS_SWITCHSTART As Long = &H14 ' Sent when beginning and ending alt-tab mode with the switch window.
'Private Const SYS_SWITCHEND As Long = &H15
'Private Const SYS_MINIMIZESTART As Long = &H16 '--window minimizing.
'Private Const SYS_MINIMIZEEND As Long = &H17 '--window about to restore.
'Private Const OB_CREATE As Long = &H8000 '-- hwnd + ID + idChild is created item
'Private Const OB_DESTROY As Long = &H8001 '-- hwnd + ID + idChild is destroyed item
'Private Const OB_SHOW As Long = &H8002 '-- hwnd + ID + idChild is shown item
'Private Const OB_HIDE As Long = &H8003 '-- hwnd + ID + idChild is hidden item
'Private Const OB_FOCUS As Long = &H8005 '-- hwnd + ID + idChild is focused item
'Private Const OB_SELECTION As Long = &H8006 '-- hwnd + ID + idChild is selected item (if only one), or idChild is OBJID_WINDOW if complex
'Private Const OB_SELECTIONADD As Long = &H8007 '-- hwnd + ID + idChild is item added
'Private Const OB_SELECTIONREMOVE As Long = &H8008 '-- hwnd + ID + idChild is item removed
'Private Const OB_SELECTIONWITHIN As Long = &H8009 '-- hwnd + ID + idChild is parent of changed selected
'Private Const OB_STATECHANGE As Long = &H800A ' hwnd + ID + idChild is item w/ state change
'Private Const OB_LOCATIONCHANGE As Long = &H800B '-- hwnd + ID + idChild is moved/sized item
'Private Const OB_NAMECHANGE As Long = &H800C '-- hwnd + ID + idChild is item w/ name change
'Private Const OB_DESCRIPTIONCHANGE As Long = &H800D '-- hwnd + ID + idChild is item w/ desc change
'Private Const OB_VALUECHANGE As Long = &H800E '-- hwnd + ID + idChild is item w/ value change
'Private Const OB_PARENTCHANGE As Long = &H800F '-- hwnd + ID + idChild is item w/ new parent
'Private Const OB_HELPCHANGE As Long = &H8010 '-- hwnd + ID + idChild is item w/ help change
'Private Const OB_DEFACTIONCHANGE As Long = &H8011 '-- hwnd + ID + idChild is item w/ def action change
'Private Const OB_ACCELERATORCHANGE As Long = &H8012 ' hwnd + ID + idChild is item w/ keybd accel change
Private Const WINEVENT_OUTOFCONTEXT As Long = 0
Private Const EVENT_MIN             As Long = 1
Private Const EVENT_MAX             As Long = 2147483647
'Private Const OB_REORDER As Long = &H8004 '-- hwnd + ID + idChild is parent of zordering children
Private Const EVENT_OB_FOREGROUND   As Long = 3
Private Const SWP_NOSIZE            As Long = 1
Private Const SWP_NOMOVE            As Long = 2
Private Const SWP_NOACTIVATE        As Long = 16
Private Const SWP_SHOWWINDOW        As Long = 64
Private Const SWP_NOOWNERZORDER     As Long = &H200
Private Const SWP_NOSENDCHANGING    As Long = &H400
Private Const HWND_BOTTOM           As Long = 1
Private Type WINEVENTPROC
    hWinEventHook As Long
    event As Long
    hWnd As Long
    idObject As Long
    idChild As Long
    idEventThread As Long
    dwmsEventTime As Long
End Type
Private Declare Function apiSetWinEventHook Lib "user32" Alias "SetWinEventHook" (ByVal eventMin As Long, ByVal eventMax As Long, ByVal hmodWinEventProc As Long, ByVal pfnWinEventProc As Long, ByVal idProcess As Long, ByVal idThread As Long, ByVal dwFlags As Long) As Long
Private Declare Function apiUnhookWinEvent Lib "user32" Alias "UnhookWinEvent" (ByVal LHandle As Long) As Long
Private Declare Function apiSetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private eHook          As Long
Public xinputToDesktop As Boolean

Public Sub StartEventHook()
    On Error Resume Next
    eHook = apiSetWinEventHook(EVENT_MIN, EVENT_MAX, 0, AddressOf WinEventFunc, 0, 0, WINEVENT_OUTOFCONTEXT)
End Sub
Public Sub StopEventHook()
    On Error Resume Next
    If eHook = 0 Then Exit Sub
    If apiUnhookWinEvent(eHook) <> 0 Then eHook = 0
End Sub
Private Function WinEventFunc(ByVal eWnd As Long, ByVal LEvent As Long, ByVal hWnd As Long, ByVal idObject As Long, ByVal idChild As Long, ByVal idEventThread As Long, ByVal dwmsEventTime As Long) As Long
    On Error Resume Next
    If LEvent = EVENT_OB_FOREGROUND Then
        If hWnd <> frmMain.hWnd Then
            Call frmMain.abort3Dxinput(False)
        End If
        Call apiSetWindowPos(frmMain.hWnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE Or SWP_NOOWNERZORDER Or SWP_NOSENDCHANGING)
    End If
End Function
