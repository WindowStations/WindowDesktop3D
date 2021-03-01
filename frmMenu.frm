VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuWallpaper 
      Caption         =   "Menu Wallpaper"
      Begin VB.Menu mnuChange 
         Caption         =   "Change Wallpaper"
      End
      Begin VB.Menu mnuDesk 
         Caption         =   "View Desktop Shortcuts"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    stFilter As String
    stCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    strFile As String
    nMaxFile As Long
    stFileTitle As String
    nMaxFileTitle As Long
    stInitialDir As String
    strTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    stDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Declare Function apiGetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (ByRef OFN As OPENFILENAME) As Long

Private Sub mnuChange_Click()
    '    On Error Resume Next
    '    Dim f    As String
    '    Dim filt As String
    '    filt = "Bitmap Files (*.bmp)" & Chr(0) & "*.bmp" & Chr(0) & "JPEG Files (*.jpg)" & Chr(0) & "*.jpg" & Chr(0) & "AVI Files (*.avi)" & Chr(0) & "*.avi" & Chr(0)
    '    f = FileOpenSave(0, filt, 1, "", "", "Select Wallpaper", -1, True, App.Path & "\images")
    '    If Trim(f) = "" Then Exit Sub
    '    If Dir(f, vbNormal) = "" Then Exit Sub
    ''    apb.StopAviCtrl
    '    If Strings.right(LCase(f), 4) = ".avi" Then
    '        frmWallpaper.Cls
    '        Set frmWallpaper.Image = Nothing
    '        frmWallpaper.Picture = Nothing
    '        Call apb.PlayAviCtrl(frmWallpaper.hWnd, f, 0, True, True, False)
    '    Else
    '        frmWallpaper.Picture = LoadPicture(f)
    '        frmWallpaper.PaintPicture frmWallpaper.Picture, 0, 0, frmWallpaper.ScaleWidth, frmWallpaper.ScaleHeight, 0, 0, frmWallpaper.Picture.Width / 26.46, frmWallpaper.Picture.Height / 26.46
    '        frmWallpaper.Picture = frmWallpaper.Image
    '    End If
    '    SaveSetting "WindowWallpaper", "Saved Wallpaper", "Path1", f
    '    Exit Sub
    'skip:
End Sub
Friend Function FileOpenSave(ByRef flags As Long, Optional ByVal Filter As String = vbNullString, Optional ByVal FilterIndex As Long = 1, Optional ByVal DefaultExt As String = vbNullString, Optional ByVal FileName As String = vbNullString, Optional ByVal DialogTitle As String = vbNullString, Optional ByVal hWnd As Long = -1, Optional ByVal OpenFile As Boolean = True, Optional ByVal inidir As String = vbNullString) As String
    On Error Resume Next
    Dim OFN         As OPENFILENAME
    Dim stFileName  As String
    Dim stFileTitle As String
    Dim fResult     As Long
    If (hWnd = -1) Then hWnd = 0
    stFileName = left(FileName & String(260, vbNullChar), 260)
    stFileTitle = String(260, vbNullChar)
    With OFN
        .lStructSize = Len(OFN)
        .hWndOwner = hWnd
        .stFilter = Filter
        .nFilterIndex = FilterIndex
        .strFile = stFileName
        .nMaxFile = Len(stFileName)
        .stFileTitle = stFileTitle
        .nMaxFileTitle = Len(stFileTitle)
        .strTitle = DialogTitle
        .flags = flags
        .stDefExt = DefaultExt
        .stInitialDir = inidir
        .hInstance = 0
        .stCustomFilter = String(260, vbNullChar)
        .nMaxCustFilter = 260
        .lpfnHook = 0
    End With
    fResult = apiGetOpenFileName(OFN)
    Dim of As OPENFILENAME
    Call apiGetOpenFileName(of)
    If fResult = 0 Then Exit Function
    flags = OFN.flags
    FileOpenSave = left(OFN.strFile, InStr(1, OFN.strFile, vbNullChar, vbBinaryCompare) - 1)
End Function

Private Sub mnuDesk_Click()
    On Error Resume Next
    Dim pth As String
    pth = CStr(CreateObject("WScript.Shell").Specialfolders("Desktop"))
    If pth = "" Then Exit Sub
    CreateObject("Wscript.Shell").Run "explorer " & pth
End Sub
