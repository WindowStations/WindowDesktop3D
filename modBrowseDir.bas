Attribute VB_Name = "modBrowseDir"
Option Explicit
Private Const BIF_STATUSTEXT        As Long = &H4&
Private Const BIF_RETURNONLYFSDIRS  As Long = 1
Private Const BIF_DONTGOBELOWDOMAIN As Long = 2
Private Const MAX_PATH              As Long = 260
Private Const WM_USER               As Long = &H400
Private Const BFFM_INITIALIZED      As Long = 1
Private Const BFFM_SELCHANGED       As Long = 2
Private Const BFFM_SETSTATUSTEXT    As Long = (WM_USER + 100)
Private Const BFFM_SETSELECTION     As Long = (WM_USER + 102)
Private Declare Function apiSendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function apiSHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolder" (ByRef lpbi As BrowseInfo) As Long
Private Declare Function apiSHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDList" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Type BrowseInfo
   hWndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type
Private m_CurrentDirectory As String
Public Function BrowseForFolder(owner As Form, Title As String, StartDir As String) As String
   On Error Resume Next
   Dim lpIDList    As Long
   Dim szTitle     As String
   Dim sBuffer     As String
   Dim tBrowseInfo As BrowseInfo
   m_CurrentDirectory = StartDir & vbNullChar
   szTitle = Title
   With tBrowseInfo
      .hWndOwner = owner.hWnd
      .lpszTitle = lstrcat(szTitle, "")
      .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_STATUSTEXT
      .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)  'get address of function.
   End With
   lpIDList = apiSHBrowseForFolder(tBrowseInfo)
   If (lpIDList) Then
      sBuffer = Space(MAX_PATH)
      apiSHGetPathFromIDList lpIDList, sBuffer
      sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
      BrowseForFolder = sBuffer
   Else
      BrowseForFolder = ""
   End If
End Function
Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
   On Error Resume Next
   Dim lpIDList As Long
   Dim ret      As Long
   Dim sBuffer  As String
   Select Case uMsg
   Case BFFM_INITIALIZED
      Call apiSendMessage(hWnd, BFFM_SETSELECTION, 1, m_CurrentDirectory)
   Case BFFM_SELCHANGED
      sBuffer = Space(MAX_PATH)
      ret = apiSHGetPathFromIDList(lp, sBuffer)
      If ret = 1 Then Call apiSendMessage(hWnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
   End Select
   BrowseCallbackProc = 0
End Function
Private Function GetAddressofFunction(add As Long) As Long
   GetAddressofFunction = add
End Function
