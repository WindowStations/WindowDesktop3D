Attribute VB_Name = "modSecureDesk"
Option Explicit
Const CCHFORMNAME                As Long = 32
Const CCHDEVICENAME              As Long = 32
Const DF_DENYOTHERACCOUNTHOOK    As Long = 0
Const DESKTOP_SECURE             As Long = 131527 'my secure version is:CREATEMENU,CREATEWINDOW,ENUMERATE,READOBJECTS,SWITCHDESKTOP,WRITEOBJECTS,READ_CONTROL
Const DESKTOP_READOBJECTS        As Long = 1 'Required to read objects on the desktop.
Const DESKTOP_SWITCHDESKTOP      As Long = 256 'Required to activate the desktop using the SwitchDesktop function.
Const INFINITE                   As Long = -1
Const SND_ASYNC                  As Long = 1
Const SND_NOSTOP                 As Long = 16
Const SND_PURGE                  As Long = 64
Const SND_FILENAME               As Long = 131072
Const SPI_SETDESKWALLPAPER       As Long = 20
Const SPIF_UPDATEINIFILE         As Long = 1
Const SPIF_SENDWININICHANGE      As Long = 2
Const UOI_NAME                   As Long = 2
Const PROCESS_QUERY_INFORMATION  As Long = &H400
Const STATUS_PENDING             As Long = &H103
Public Const DESKTOP_LOGON       As String = "Winlogon"
Public Const DESKTOP_WINSTATION0 As String = "WinSta0"
Public Const DESKTOP_DEFAULT     As String = "Default"
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
Private Type STARTUPINFOW
    cbSize As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Private Declare Function apiCloseDesktop Lib "user32" Alias "CloseDesktop" (ByVal hDesktop As Long) As Long
Private Declare Function apiCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long
Private Declare Function apiCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDst As String, ByRef pSrc As Long, ByVal ByteLen As Long) As Long
Private Declare Function apiCreateDesktop Lib "user32" Alias "CreateDesktopW" (ByVal lpszDesktop As Long, ByVal lpszDevice As Long, ByRef pDevmode As Long, ByVal dwFlags As Long, ByVal dwDesiredAccess As Long, ByRef lpsa As Long) As Long
Private Declare Function apiCreateProcess Lib "kernel32" Alias "CreateProcessW" (ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, ByRef lpProcessAttributes As Long, ByRef lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByRef lpEnvironment As Long, ByVal lpCurrentDirectory As Long, ByRef lpStartupInfo As STARTUPINFOW, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function apiEnumDesktops Lib "user32" Alias "EnumDesktopsA" (ByVal hSta As Long, ByVal lEnumFunc As Long, ByVal lParam As Long) As Long
'Private Declare Function apiFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Boolean
Private Declare Function apiGetCurrentThreadId Lib "kernel32" Alias "GetCurrentThreadId" () As Long
Private Declare Function apiGetDC Lib "user32" Alias "GetDC" (ByVal hwnd As Long) As Long
Private Declare Function apiGetExitCodeProcess Lib "kernel32" Alias "GetExitCodeProcess" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
Private Declare Function apiGetProcessWindowStation Lib "user32" Alias "GetProcessWindowStation" () As Long
Private Declare Function apiGetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function apiGetThreadDesktop Lib "user32" Alias "GetThreadDesktop" (ByVal dwThread As Long) As Long
Private Declare Function apiGetUserObjectInformation Lib "user32" Alias "GetUserObjectInformationA" (ByVal hObj As Long, ByVal nIndex As Long, ByVal pvInfo As String, ByVal nLength As Long, ByRef lpnLengthNeeded As Long) As Long
Private Declare Function apiGetWindowDC Lib "user32" Alias "GetWindowDC" (ByVal hwnd As Long) As Long
Private Declare Function apiOpenInputDesktop Lib "user32" Alias "OpenInputDesktop" (ByVal dwFlags As Long, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Private Declare Function apiPaintDesktop Lib "user32" Alias "PaintDesktop" (ByVal hDC As Long) As Long
Private Declare Function apiPlaySound Lib "winmm" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function apiSetThreadDesktop Lib "user32" Alias "SetThreadDesktop" (ByVal hDesktop As Long) As Long
Private Declare Function apiStrLen Lib "kernel32" Alias "lstrlenA" (ByVal lString As Long) As Long
Private Declare Function apiSwitchDesktop Lib "user32" Alias "SwitchDesktop" (ByVal hDesktop As Long) As Long
Private Declare Function apiSystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lParam As String, ByVal fuWinIni As Long) As Long
Private Declare Function apiWaitForSingleObject Lib "kernel32" Alias "WaitForSingleObject" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function apiOpenProcess Lib "kernel32" Alias "OpenProcess" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function apiOpenDesktop Lib "user32" Alias "OpenDesktopA" (ByVal lpszDesktop As String, ByVal dwFlags As Long, ByVal fInherit As Long, ByVal dwDesiredAccess As Long) As Long

'Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function apiSleepEx Lib "kernel32" Alias "SleepEx" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Public DESKTOP_X     As String
Private newDskTop    As String
Private oldDskThread As Long
Private oldDskInput  As Long
Private hwnDsk       As Long
Public lDesktops     As String

'Public Function IsDiskDrivePresent() As Boolean
'    On Error Resume Next
'    Dim strComputer   As String
'    Dim objWMIService As Object
'    Dim colitems      As Object
'    Dim objitem       As Object
'    Dim s         As String
'    strComputer = "."
'    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
'    Set colitems = objWMIService.ExecQuery("Select * from Win32_DiskDrive")
'    For Each objitem In colitems
'       s = ""
'       s = objitem.SerialNumber
'       If Trim(s) = USB_SERIAL_NUMBER Then
'         IsDiskDrivePresent = True
'         Exit For
'       End If
'    Next
'    Set objWMIService = Nothing
'    Set colitems = Nothing
'    Set objitem = Nothing
'End Function
'Public Function IsDiskDrivePresent2() As Boolean
'    On Error Resume Next
'    Dim d       As String
'    Dim I       As Long
'    Dim s       As String
'    Dim ret     As Long
'    Dim keer    As Long
'    Dim strSave As String
'    With New clsUsbFlashSerial
'        strSave = String(260, Chr(0))
'        ret = GetLogicalDriveStrings(260, strSave)
'        For keer = 1 To 100
'            If Left(strSave, InStr(1, strSave, Chr(0))) = Chr(0) Then Exit For
'            d = LCase(Left(strSave, InStr(1, strSave, Chr(0)) - 2))
'            s = .Lookup(d)
'            If Len(s) < 1 Then s = ""
'            If Trim$(s) <> "" Then
'                If Trim$(s) = USB_SERIAL_NUMBER Then
'                    IsDiskDrivePresent2 = True
'                    Exit For
'                End If
'            End If
'            strSave = Right(strSave, Len(strSave) - InStr(1, strSave, Chr(0)))
'        Next
'    End With
'End Function
'Public Function PromptSecure(ByVal message As String, ByVal title As String, Optional ByVal timeout As Long, Optional ByVal exepath As String) As Long    '(ByVal message As String, Optional ByVal title As String, Optional ByVal dskname As String, Optional ByVal exepath As String)
'    On Error Resume Next
'    Dim dskname As String
'    Dim rn      As Long
'    Randomize
'    rn = Rnd * (2147483647 - 1) + 1
'    dskname = CStr(rn) 'Set desk name to random string
'
'    dskname = "AA010728150356283255"
'
'    oldDskThread = apiGetThreadDesktop(apiGetCurrentThreadId)
'    oldDskInput = apiOpenInputDesktop(0, False, DESKTOP_SWITCHDESKTOP)
'    If CreateDesktop(dskname) = 0 Then Exit Function
'    Call PlaySnd("") 'play classic UAC sound or xp log off
'    SwitchToDeskTop
'    Call StartProcess(exepath) 'synchro- waits until process is finished.  Optional application to start
'    KillCTFMON 'kill off extra process started by Windows
''    CloseDeskTop 'Close the desktop we created
''    apiSetThreadDesktop (oldDskThread) 'Set the thread desktop back
''    apiSwitchDesktop (oldDskInput) 'If switched clear old desk
'End Function
'Public Function CreateDesktop(ByVal sDesktopName As String) As Long
'    On Error Resume Next
'    hwnDsk = apiCreateDesktop(StrPtr(sDesktopName), ByVal 0, ByVal 0, 0, DESKTOP_SECURE, ByVal 0)
'    If hwnDsk = 0 Then CreateDesktop = 0: Exit Function
'    newDskTop = sDesktopName: CreateDesktop = hwnDsk
'End Function
'Public Function SwitchToDeskTop() As Long
'    On Error Resume Next
'    Dim st As Long
'    Dim sd As Long
'    st = apiSetThreadDesktop(hwnDsk)
'    sd = apiSwitchDesktop(hwnDsk)
'    If sd <> 0 Then SwitchToDeskTop = 1
'End Function
Public Function SwitchToDefaultDeskTop() As Long
    On Error Resume Next
    Dim st As Long
    Dim sd As Long
    hwnDsk = apiOpenDesktop(lpszDesktop:="Default", dwFlags:=0, fInherit:=False, dwDesiredAccess:=DESKTOP_SWITCHDESKTOP)
    st = apiSetThreadDesktop(hwnDsk)
    sd = apiSwitchDesktop(hwnDsk)
End Function
'Public Function StartProcess(ByVal exepath As String) '(ByVal sPath As String, ByVal message As String) As Long
'    On Error Resume Next
'    Dim psi   As STARTUPINFOW
'    Dim pInfo As PROCESS_INFORMATION
'    psi.cbSize = Len(psi)
'    psi.lpTitle = StrPtr(newDskTop)
'    psi.lpDesktop = StrPtr(newDskTop)
'    'Call apiCreateProcess(StrPtr(App.Path & "\DesktopTaskbar.exe"), ByVal 0, ByVal 0, ByVal 0, 1, 0, ByVal 0, ByVal 0, psi, pInfo)
'    StartProcess = apiCreateProcess(StrPtr(exepath), ByVal 0, ByVal 0, ByVal 0, 1, 0, ByVal 0, ByVal 0, psi, pInfo)
''    If StartProcess <> 0 Then
''        Call apiWaitForSingleObject(pInfo.hProcess, INFINITE)   'Wait until the process has completed
''        apiCloseHandle (pInfo.hProcess)
''        apiCloseHandle (pInfo.hThread)
''    End If
'End Function
'Private Sub KillCTFMON()
'    Dim objshell
'    Set objshell = CreateObject("Wscript.Shell")
'    objshell.Run "taskkill /IM ctfmon.exe", 0, True
'End Sub
'Private Function GetExitCode(ByVal hProcess As Long) As Long
'    If hProcess = 0 Then GetExitCode = -1: Exit Function
'    Dim exitCode As Long
'    Dim I        As Long
'    For I = 1 To 32767
'        Call apiGetExitCodeProcess(hProcess, exitCode)
'        DoEvents
'        If exitCode <> STATUS_PENDING Then Exit For
'    Next
'    GetExitCode = exitCode
'End Function
'Public Sub CloseDeskTop()
'    On Error Resume Next
'    apiCloseDesktop (hwnDsk)
'End Sub
''Public Sub PlaySnd(Optional ByVal uacPath As String)
''    On Error Resume Next
''    'Get path to windows media folder, for stock UAC sound
''    Dim medPath As String
''    medPath = GetSystemDirectory
''    medPath = Left(medPath, Len(medPath) - 9)
''    medPath = medPath & "\media\"
''    If uacPath = "on" Then
''        If Dir(medPath & "Windows User Account Control.wav") = "" Then
''            uacPath = medPath & "Windows XP Logon Sound.wav"
''        Else
''            uacPath = medPath & "Windows User Account Control.wav"
''        End If
''    Else
''        If Dir(medPath & "Windows User Account Control.wav") = "" Then
''            uacPath = medPath & "Windows XP Logoff Sound.wav"
''        Else
''            uacPath = medPath & "Windows User Account Control.wav"
''        End If
''    End If
''    'Clear sound, and then play
''    Call apiPlaySound(vbNullString, 0, SND_FILENAME Or SND_ASYNC)
''    Call apiPlaySound(uacPath, 0, SND_FILENAME Or SND_ASYNC)
''End Sub
'Public Function GetSystemDirectory() As String
'    On Error Resume Next
'    Dim ret As Long
'    GetSystemDirectory = Space(260)    'Create a buffer
'    ret = apiGetSystemDirectory(GetSystemDirectory, 260)  'Get sysdir
'    GetSystemDirectory = Left(GetSystemDirectory, ret) 'Remove chr$(0)'s
'End Function
'Public Function SetDesktopWallpaper(Optional ByVal imgPath As String = "") As Long
'    On Error Resume Next
'    SetDesktopWallpaper = apiSystemParametersInfo(SPI_SETDESKWALLPAPER, 0, imgPath, SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE)
'End Function
'Public Function PaintDesktop(ByVal hwnd As Long) As Long
'    On Error Resume Next
'    PaintDesktop = apiPaintDesktop(apiGetWindowDC(hwnd))
'End Function
Public Function GetDesktopName() As String
    On Error Resume Next
    DESKTOP_X = DESKTOP_DEFAULT
    Dim hDesktop As Long
    hDesktop = apiOpenInputDesktop(DF_DENYOTHERACCOUNTHOOK, False, DESKTOP_READOBJECTS)
    If hDesktop = 0 Then GetDesktopName = "": Exit Function 'exit if desktop cannot open
    Dim uInf  As Long
    Dim lSize As Long
    Dim lLen  As Long
    Dim bf    As String
    GetDesktopName = "" 'Initialize to default
    lSize = (Len(DESKTOP_X) + 10) * 2
    bf = String(lSize - 1, Chr(0))  'buffer
    uInf = apiGetUserObjectInformation(hDesktop, UOI_NAME, bf, lSize, lLen)
    If uInf <> 0 Then 'If function failed no sense stripping buffer
        Dim iPos As Long
        iPos = InStr(bf, Chr(0)) '+ 1 'bf.IndexOf(Chr(0)) + 1
        If iPos > 1 Then bf = Left(bf, iPos - 1)
        GetDesktopName = bf
    End If
    Call apiCloseHandle(hDesktop)
End Function
Public Function GetWindowStationName(ByVal wshwnd As Long) As String
    On Error Resume Next
    'DESKTOP_X = DESKTOP_DEFAULT
   ' Dim hDesktop As Long
    'hDesktop = apiOpenInputDesktop(DF_DENYOTHERACCOUNTHOOK, False, DESKTOP_READOBJECTS)
   ' If hDesktop = 0 Then GetDesktopName = "": Exit Function 'exit if desktop cannot open
    Dim uInf  As Long
    Dim lSize As Long
    Dim lLen  As Long
    Dim bf    As String
    GetWindowStationName = "" 'Initialize to default
    lSize = (Len(DESKTOP_X) + 10) * 2
    bf = String(lSize - 1, Chr(0))  'buffer
    uInf = apiGetUserObjectInformation(wshwnd, UOI_NAME, bf, lSize, lLen)
    If uInf <> 0 Then 'If function failed no sense stripping buffer
        Dim iPos As Long
        iPos = InStr(bf, Chr(0)) '+ 1 'bf.IndexOf(Chr(0)) + 1
        If iPos > 1 Then bf = Left(bf, iPos - 1)
        GetWindowStationName = bf
    End If
    apiCloseHandle (wshwnd)
End Function
Public Function GetDesktops() As String
    On Error Resume Next
    Call apiEnumDesktops(apiGetProcessWindowStation, AddressOf EnumDesktopProc, 0)
    GetDesktops = lDesktops
End Function
Private Function EnumDesktopProc(ByVal lDesktop As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Dim Buffer As String
    Buffer = Space(apiStrLen(lDesktop)) 'Create buffer
    'Call copy memory correctly with ByVal in ByRef params/Uhgg
    Call apiCopyMemory(ByVal Buffer, ByVal lDesktop, Len(Buffer))
    lDesktops = lDesktops & Buffer & vbCrLf 'Append string of desktops
    EnumDesktopProc = 1 'Return complete
End Function
'Const DF_ALLOWOTHERACCOUNTHOOK As Long = 1
'Const DESKTOP_CREATEWINDOW As Long = 2 'Required to create a window on the desktop.
'Const DESKTOP_CREATEMENU As Long = 4 'Required to create a menu on the desktop.
'Const DESKTOP_HOOKCONTROL As Long = 8 'Required to establish any of the window hooks.
'Const DESKTOP_JOURNALRECORD As Long = 16 ''Required to perform journal recording on a desktop.
'Const DESKTOP_JOURNALPLAYBACK As Long = 32 'Required to perform journal playback on a desktop.
'Const DESKTOP_ENUMERATE As Long = 64 'Required for the desktop to be enumerated.
'Const DESKTOP_WRITEOBJECTS As Long = 128 'Required to write objects on the desktop.
'Const READ_CONTROL As Long = 131072 'STANARD_
'Const GENERIC_READ As Long = 131137 'enumerate,readobject,read control
'Const GENERIC_WRITE As Long = 131262 'createmenu,createwindow,hookcontrol,journalrecord,journalplayback,writeobjects
'Const GENERIC_EXECUTE As Long = 131328 'switchdesktop,read control
'Const GENERIC_ALL As Long = 268435456 'old all all?
'Const GENERIC_ALL2 As Long = 131583 'new all?
