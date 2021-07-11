VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDesktopShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
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
Private Const MAX_PATH                      As Integer = 260
Private Const S_OK = &H0
Private Const S_FALSE = &H1
Private Const E_INVALIDARG = &H80070057
Private Const SHGFP_TYPE_CURRENT = 0
Private Const SHGFP_TYPE_DEFAULT = 1
'Private Const CSIDL_FLAG_CREATE = &H8000
'Private Const CSIDL_ADMINTOOLS = &H30
'Private Const CSIDL_ALTSTARTUP = &H1D
'Private Const CSIDL_APPDATA = &H1A
'Private Const CSIDL_BITBUCKET = &HA
'Private Const CSIDL_CDBURN_AREA = &H3B
'Private Const CSIDL_COMMON_ADMINTOOLS = &H2F
'Private Const CSIDL_COMMON_ALTSTARTUP = &H1E
'Private Const CSIDL_COMMON_APPDATA = &H23
'Private Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
'Private Const CSIDL_COMMON_DOCUMENTS = &H2E
'Private Const CSIDL_COMMON_FAVORITES = &H1F
'Private Const CSIDL_COMMON_MUSIC = &H35
'Private Const CSIDL_COMMON_PICTURES = &H36
'Private Const CSIDL_COMMON_PROGRAMS = &H17
'Private Const CSIDL_COMMON_STARTMENU = &H16
'Private Const CSIDL_COMMON_STARTUP = &H18
'Private Const CSIDL_COMMON_TEMPLATES = &H2D
'Private Const CSIDL_COMMON_VIDEO = &H37
'Private Const CSIDL_CONTROLS = &H3
'Private Const CSIDL_COOKIES = &H21
'Private Const CSIDL_DESKTOP = &H0
'Private Const CSIDL_DESKTOPDIRECTORY = &H10
'Private Const CSIDL_DRIVES = &H11
'Private Const CSIDL_FAVORITES = &H6
'Private Const CSIDL_FONTS = &H14
'Private Const CSIDL_HISTORY = &H22
'Private Const CSIDL_INTERNET = &H1
'Private Const CSIDL_INTERNET_CACHE = &H20
'Private Const CSIDL_LOCAL_APPDATA = &H1C
'Private Const CSIDL_MYDOCUMENTS = &HC
'Private Const CSIDL_MYMUSIC = &HD
'Private Const CSIDL_MYPICTURES = &H27
'Private Const CSIDL_MYVIDEO = &HE
'Private Const CSIDL_NETHOOD = &H13
'Private Const CSIDL_NETWORK = &H12
'Private Const CSIDL_PERSONAL = &H5
'Private Const CSIDL_PRINTERS = &H4
'Private Const CSIDL_PRINTHOOD = &H1B
'Private Const CSIDL_PROFILE = &H28
'Private Const CSIDL_PROFILES = &H3E
'Private Const CSIDL_PROGRAM_FILES = &H26
'Private Const CSIDL_PROGRAM_FILES_COMMON = &H2B
'Private Const CSIDL_PROGRAMS = &H2
'Private Const CSIDL_RECENT = &H8
'Private Const CSIDL_SENDTO = &H9
'Private Const CSIDL_STARTMENU = &HB
'Private Const CSIDL_STARTUP = &H7
'Private Const CSIDL_SYSTEM = &H25
'Private Const CSIDL_TEMPLATES = &H15
'Private Const CSIDL_WINDOWS = &H24
Private Type SHITEMID
   cb As Long
   abID As Byte
End Type
Private Type ITEMIDLIST
   mkid As SHITEMID
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
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetFolderPath Lib "shell32.dll" Alias "SHGetFolderPathA" (ByVal hWndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal lpszPath As String) As Long
Private Declare Function apiShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Private Declare Function apiCreateProcess Lib "kernel32" Alias "CreateProcessW" (ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, ByRef lpProcessAttributes As Long, ByRef lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByRef lpEnvironment As Long, ByVal lpCurrentDirectory As Long, ByRef lpStartupInfo As STARTUPINFOW, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (ByRef pPictDesc As PICTDESC, ByRef RefIID As GUID, ByVal fPictureOwnsHandle As Long, ByRef ppvObj As StdPicture) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Friend Function fGetSpecialFolder(CSIDL As Long) As String
   Dim sPath As String
   Dim IDL   As ITEMIDLIST
   fGetSpecialFolder = ""
   If SHGetSpecialFolderLocation(0, CSIDL, IDL) = 0 Then
      sPath = VBA.Space$(MAX_PATH)
      If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath) Then
         fGetSpecialFolder = VBA.left$(sPath, InStr(sPath, vbNullChar) - 1) & ""
      End If
   End If
End Function
Friend Function GetSpecialFolder2(ByVal lngFolder As Long) As String
   Dim strBuffer As String * 1000
   Dim strPath   As String
   Dim lngReturn As Long
   lngReturn = SHGetFolderPath(0&, lngFolder, 0&, SHGFP_TYPE_CURRENT, strBuffer)
   If lngReturn = S_OK Then
      strPath = VBA.left$(strBuffer, InStr(strBuffer, VBA.Chr$(0)) - 1)
   Else
      strPath = "(error)"
   End If
   GetSpecialFolder2 = strPath
End Function
Private Function GetTarget(ByVal FileName As String) As String
   On Error Resume Next
   Dim obj      As Object
   Dim Shortcut As Object
   Set obj = CreateObject("WScript.Shell")
   Set Shortcut = obj.CreateShortcut(FileName)
   GetTarget = VBA.LCase(Shortcut.TargetPath)
   Set obj = Nothing
   Set Shortcut = Nothing
End Function
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
   Const FILE_ATTRIBUTE_READONLY  As Long = &H1
   Const FILE_ATTRIBUTE_HIDDEN    As Long = &H2
   Const FILE_ATTRIBUTE_SYSTEM    As Long = &H4
   Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
   Const FILE_ATTRIBUTE_ARCHIVE   As Long = &H20
   Const FILE_ATTRIBUTE_NORMAL    As Long = &H80
   Const FILE_ATTRIBUTE_TEMPORARY As Long = &H100
   Call SHGetFileInfo(sFileName, 0, tFileInfo, Len(tFileInfo), SHGFI_DISPLAYNAME Or SHGFI_TYPENAME Or SHGFI_LARGEICON Or SHGFI_ICON)
   hIcon = tFileInfo.hIcon
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
Friend Function GetDesktopShellPaths(ByVal pth As String) As String()
   On Error Resume Next
   Dim s() As String
   Dim fso As Object
   ReDim Preserve s(0)
   Set fso = CreateObject("Scripting.FileSystemObject")
   If Not fso Is Nothing Then
      Dim fld As Object
      Dim fl  As Object
      Dim fil As Variant
      Dim z   As Long
      Set fld = fso.GetFolder(pth)
      For Each fl In fld.SubFolders
         ReDim Preserve s(z)
         s(z) = fl
         z = z + 1
      Next
      For Each fil In fld.Files
         If VBA.LCase(VBA.right(fil, 4)) <> ".ini" Then
            ReDim Preserve s(z)
            s(z) = fil
            z = z + 1
         End If
      Next
      GetDesktopShellPaths = s
      Set fil = Nothing
      Set fld = Nothing
      Set fl = Nothing
      Set fso = Nothing
   End If
End Function
