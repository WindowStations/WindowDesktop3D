Attribute VB_Name = "modKeyHook"
Const WH_KEYBOARD_LL As Long = 13
Const HC_GETNEXT     As Long = 1
Const WM_SYSKEYDOWN  As Long = 260
Const WM_SYSKEYUP    As Long = 261
Const WM_KEYDOWN     As Long = 256
Const WM_KEYUP       As Long = 257
Private Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type
Private Declare Function apiSetWindowsKeyHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function apiUnhookWindowsHookEx Lib "user32" Alias "UnhookWindowsHookEx" (ByVal hHook As Long) As Long
Private Declare Function apiCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As KBDLLHOOKSTRUCT, ByVal pSource As Long, ByVal cb As Long) As Long
Private Declare Function apiCallNextKeyHookEx Lib "user32" Alias "CallNextHookEx" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private hKey        As Long
Private hMouse      As Long
Public keymovespeed As Double

Public Sub HookKeyboard()
    If hKey <> 0 Then Exit Sub
    hKey = apiSetWindowsKeyHookEx(WH_KEYBOARD_LL, AddressOf Callback, App.hInstance, 0)
End Sub
Public Sub UnhookKeyboard()
    On Error Resume Next
    If hKey = 0 Then Exit Sub
    If apiUnhookWindowsHookEx(hKey) <> 0 Then hKey = 0
End Sub

Private Function Callback(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Static hStruct As KBDLLHOOKSTRUCT
    Call apiCopyMemory(hStruct, lParam, Len(hStruct))
    If wParam = WM_KEYDOWN Or wParam = WM_SYSKEYDOWN Then
        keymovespeed = 0
        If hStruct.vkCode = vbKeyW Then frmMain.khclass.KeyWDown: GoTo skip
        If hStruct.vkCode = vbKeyA Then frmMain.khclass.KeyADown: GoTo skip
        If hStruct.vkCode = vbKeyS Then frmMain.khclass.KeySDown: GoTo skip
        If hStruct.vkCode = vbKeyD Then frmMain.khclass.KeyDDown: GoTo skip
        If hStruct.vkCode = vbKeyEscape Then frmMain.khclass.KeyEscDown: GoTo skip
        '        If hStruct.vkCode = vbKey4 And hStruct.flags = 32 Then
        '            frmMain.khclass.KeyAltF4Down
        '            Callback = 1
        '            Return
        '        End If
skip:
    Else
        If hStruct.vkCode = vbKeyW Then frmMain.khclass.KeyWUp: GoTo skip2
        If hStruct.vkCode = vbKeyA Then frmMain.khclass.KeyAUp: GoTo skip2
        If hStruct.vkCode = vbKeyS Then frmMain.khclass.KeySUp: GoTo skip2
        If hStruct.vkCode = vbKeyD Then frmMain.khclass.KeyDUp: GoTo skip2
        If hStruct.vkCode = vbKeyEscape Then frmMain.khclass.KeyEscUp: GoTo skip2
skip2:
        keymovespeed = 0
    End If
    Callback = apiCallNextKeyHookEx(hKey, Code, wParam, lParam) 'Call next key hook if no action
End Function
