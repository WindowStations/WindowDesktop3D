VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcessWTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const WTS_CURRENT_SERVER_HANDLE As Long = 0
Private Type WTS_PROCESS_INFO
    SessionID As Long
    ProcessID As Long
    pProcessName As Long
    pUserSid As Long
End Type
Private Declare Function apiLookupAccountSid Lib "advapi32" Alias "LookupAccountSidA" (ByVal lpSystemName As String, ByVal sid As Long, ByVal name As String, ByRef cbName As Long, ByVal ReferencedDomainName As String, ByRef cbReferencedDomainName As Long, ByRef peUse As Integer) As Long
Private Declare Function apiWTSEnumerateProcesses Lib "wtsapi32" Alias "WTSEnumerateProcessesA" (ByVal hServer As Long, ByVal Reserved As Long, ByVal Version As Long, ByRef ppProcessInfo As Long, ByRef pCount As Long) As Long
Private Declare Sub apiWTSFreeMemory Lib "wtsapi32" Alias "WTSFreeMemory" (ByVal pMemory As Long)
Private Declare Sub apiCopyMemoryWTSLONG Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As WTS_PROCESS_INFO, ByRef Source As Long, ByVal Length As Long)
Private Declare Sub apiCopyMemoryBYTELONG Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Byte, ByRef Source As Long, ByVal Length As Long)

'Friend Function GetWTSProcesses(ByRef lvis As ListItems)
'    On Error Resume Next
'    Dim RetVal As Long
'    Dim Count As Long
'    Dim i As Integer
'    Dim lpBuffer As Long
'    Dim p As Long
'    Dim udtProcessInfo As WTS_PROCESS_INFO
'    Dim itmAdd As ListItem
'    lvis.Clear
'    RetVal = apiWTSEnumerateProcesses(WTS_CURRENT_SERVER_HANDLE, 0&, 1, lpBuffer, Count)
'    If RetVal <> 0 Then
'        p = lpBuffer
'        Do
'            i = i + 1
'            Call apiCopyMemoryWTSLONG(udtProcessInfo, ByVal p, LenB(udtProcessInfo))
'            Dim pname As String
'            Dim uname As String
'            Dim pid As Long
'            Dim sid As Long
'            pid = udtProcessInfo.ProcessID
'            sid = udtProcessInfo.SessionID
'            If udtProcessInfo.ProcessID = 0 Then
'                pname = "System Idle Process"
'            Else
'                pname = GetStringFromLP(udtProcessInfo.pProcessName)
'            End If
'            uname = GetUserName(udtProcessInfo.pUserSid)
'            Set itmAdd = lvis.add(i, , CStr(sid))
'            itmAdd.SubItems(1) = CStr(pid)
'            If udtProcessInfo.ProcessID = 0 Then
'              itmAdd.SubItems(2) = "System Idle Process"
'            Else
'               itmAdd.SubItems(2) = pname
'            End If
'            itmAdd.SubItems(3) = uname
'            p = p + LenB(udtProcessInfo) ' Increment to next WTS_PROCESS_INO structure in the buffer
'            If i = Count Then Exit Do
'        Loop
'        Set itmAdd = Nothing
'        Call apiWTSFreeMemory(lpBuffer)    'Free your memory buffer
'    End If
'    Set GetWTSProcesses = lvis
'End Function
Friend Function GetWTSPIDs() As Long()
    On Error Resume Next
    Dim pids() As Long
    Dim RetVal As Long
    Dim Count As Long
    Dim i As Integer
    Dim lpBuffer As Long
    Dim p As Long
    Dim udtProcessInfo As WTS_PROCESS_INFO
    RetVal = apiWTSEnumerateProcesses(WTS_CURRENT_SERVER_HANDLE, 0&, 1, lpBuffer, Count)
    If RetVal <> 0 Then
        p = lpBuffer
        Do
            i = i + 1
            Call apiCopyMemoryWTSLONG(udtProcessInfo, ByVal p, LenB(udtProcessInfo))
            Dim pid As Long
            pid = udtProcessInfo.ProcessID
            ReDim Preserve pids(i)
            pids(i) = pid
            p = p + LenB(udtProcessInfo)
            If i = Count Then Exit Do
        Loop
        Call apiWTSFreeMemory(lpBuffer)
    End If
    GetWTSPIDs = pids
End Function
Friend Function GetWTSPID(ByVal processname As String) As Long
    On Error Resume Next
    Dim RetVal As Long
    Dim Count As Long
    Dim i As Integer
    Dim lpBuffer As Long
    Dim p As Long
    Dim udtProcessInfo As WTS_PROCESS_INFO
    Dim pidret As Long
    RetVal = apiWTSEnumerateProcesses(WTS_CURRENT_SERVER_HANDLE, 0&, 1, lpBuffer, Count)
    If RetVal <> 0 Then
        p = lpBuffer
        Do
            i = i + 1
            Call apiCopyMemoryWTSLONG(udtProcessInfo, ByVal p, LenB(udtProcessInfo))
            Dim pname As String
            Dim pid As Long
            pid = udtProcessInfo.ProcessID
            If udtProcessInfo.ProcessID = 0 Then
                pname = "System Idle Process"
            Else
                pname = GetStringFromLP(udtProcessInfo.pProcessName)
            End If
            If LCase(pname) = LCase(processname) Then pidret = pid
            p = p + LenB(udtProcessInfo) ' Increment to next WTS_PROCESS_INO structure in the buffer
            If i = Count Then Exit Do
            If pidret <> 0 Then Exit Do
        Loop
        Call apiWTSFreeMemory(lpBuffer)
    End If
    GetWTSPID = pidret
End Function
Function GetUserName(sid As Long) As String
    On Error Resume Next
    Dim retname As String
    Dim retdomain As String
    retname = String(255, 0)
    retdomain = String(255, 0)
    Call apiLookupAccountSid(vbNullString, sid, retname, 255, retdomain, 255, 0)
    GetUserName = left(retdomain, InStr(retdomain, vbNullChar) - 1) & "\" & left$(retname, InStr(retname, vbNullChar) - 1)
End Function
Private Function GetStringFromLP(ByVal StrPtr As Long) As String
    Dim b As Byte
    Dim tempStr As String
    Dim bufferStr As String
    Do
        Call apiCopyMemoryBYTELONG(b, ByVal StrPtr, 1)        ' Get the byte/character that StrPtr is pointing to.
        If b = 0 Then Exit Do  ' If you've found a null character, then you're done.
        tempStr = Chr(b)  ' Get the character for the byte's value
        bufferStr = bufferStr & tempStr 'Add it to the string
        StrPtr = StrPtr + 1  ' Increment the pointer to next byte/char
    Loop
    GetStringFromLP = bufferStr
End Function
