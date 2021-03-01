Attribute VB_Name = "modEnumWindows"
Option Explicit
Private Declare Function apiEnumWindows Lib "user32" Alias "EnumWindows" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function apiEnumChildWindows Lib "user32" Alias "EnumChildWindows" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function apiGetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function apiGetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function apiGetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function apiFindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function apiFindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function apiIsWindow Lib "user32" Alias "IsWindow" (ByVal hWnd As Long) As Long
Private cwnds()     As Long
Private handlecount As Long
Public Function ChildWindows(Optional ByVal hWnd As Long) As Long() 'If hWnd = 0 it returns all the top-level windows.
    handlecount = 0
    If hWnd = 0 Then apiEnumWindows AddressOf EnumWindowsProc, 1
    If hWnd <> 0 Then apiEnumChildWindows hWnd, AddressOf EnumWindowsProc, 1
    ReDim Preserve cwnds(handlecount) As Long
    ChildWindows = cwnds()
End Function
Private Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    If apiIsWindow(hWnd) <> 0 Then
        If handlecount = 0 Then ReDim cwnds(100) As Long
        If handlecount >= UBound(cwnds) Then ReDim Preserve cwnds(handlecount + 100) As Long
        handlecount = handlecount + 1
        cwnds(handlecount) = hWnd
    End If
    EnumWindowsProc = 1
End Function

Public Function GetClassName(ByVal hWnd As Long) As String
    On Error Resume Next
    Dim tLength As Long
    Dim rValue  As Long
    Dim css     As String
    css = "" ''''''''''''''''''''''Initialize string for class name
    css = Strings.Space(260) ''Pad with buffer
    rValue = apiGetClassName(hWnd, css, 260) 'Get classname
    css = left(css, rValue) 'Strip buffer
    GetClassName = css
End Function
Public Function GetWindowText(ByVal hWnd As Long) As String
    On Error Resume Next
    Dim tLength As Long
    Dim rValue  As Long
    Dim txt     As String
    tLength = apiGetWindowTextLength(hWnd) + 4 'Get length
    txt = Strings.Space(tLength) 'Pad with buffer
    rValue = apiGetWindowText(hWnd, txt, tLength) 'Get text
    txt = left(txt, rValue) 'Strip buffer
End Function
