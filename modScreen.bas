Attribute VB_Name = "modScreen"
Option Explicit
Private Const CCDEVICENAME             As Long = 32
Private Const CCFORMNAME               As Long = 32
Private Const DM_BITSPERPEL            As Long = &H40000
Private Const DM_PELSWIDTH             As Long = &H80000
Private Const DM_PELSHEIGHT            As Long = &H100000
Private Const CDS_UPDATEREGISTRY       As Long = &H1
Private Const CDS_TEST                 As Long = &H4
Private Const DISP_CHANGE_SUCCESSFUL   As Long = 0
Private Const DISP_CHANGE_RESTART      As Long = 1
Private Const MONITORINFOF_PRIMARY     As Long = &H1
Private Const MONITOR_DEFAULTTONEAREST As Long = &H2
Private Const MONITOR_DEFAULTTONULL    As Long = &H0
Private Const MONITOR_DEFAULTTOPRIMARY As Long = &H1
Public Type DISPLAY_DEVICE
   cb As Long
   DeviceName As String * 32
   DeviceString As String * 128
   StateFlags As Long
   DeviceID As String * 128
   DeviceKey As String * 128
End Type
Private Type RECT
   left            As Long
   top             As Long
   right           As Long
   bottom          As Long
End Type
Private Type MONITORINFO
   cbSize          As Long
   rcMonitor       As RECT
   rcWork          As RECT
   dwFlags         As Long
End Type
Private Type typDevMODE
   dmDeviceName As String * CCDEVICENAME
   dmSpecVersion As Integer
   dmDriverVersion As Integer
   dmSize As Integer
   dmDriverExtra As Integer
   dmFields As Long
   dmOrientation As Integer
   dmPaperSize As Integer
   dmPaperLength As Integer
   dmPaperWidth As Integer
   dmScale As Integer
   dmCopies As Integer
   dmDefaultSource As Integer
   dmPrintQuality As Integer
   dmColor As Integer
   dmDuplex As Integer
   dmYResolution As Integer
   dmTTOption As Integer
   dmCollate As Integer
   dmFormName As String * CCFORMNAME
   dmUnusedPadding As Integer
   dmBitsPerPel As Integer
   dmPelsWidth As Long
   dmPelsHeight As Long
   dmDisplayFlags As Long
   dmDisplayFrequency As Long
End Type
Private Declare Function apiGetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
Private Declare Function apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function apiUnionRect Lib "user32" Alias "UnionRect" (ByRef lprcDst As RECT, ByRef lprcSrc1 As RECT, ByRef lprcSrc2 As RECT) As Long
Private Declare Function apiMonitorFromPoint Lib "user32" Alias "MonitorFromPoint" (ByVal x As Long, ByVal y As Long, ByVal dwFlags As Long) As Long
Private Declare Function apiMonitorFromRect Lib "user32" Alias "MonitorFromRect" (ByRef lprc As RECT, ByVal dwFlags As Long) As Long
Private Declare Function apiMonitorFromWindow Lib "user32" Alias "MonitorFromWindow" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long
Public screens() As Screen
Public hMon      As Long
Public Function MonitorEnumProc(ByVal hMonitor As Long, ByVal hdcMonitor As Long, ByRef lprcMonitor As RECT, ByRef dwData As Long) As Long
   On Error Resume Next
   Dim rects() As RECT
   ReDim Preserve rects(dwData)
   ReDim Preserve screens(dwData)
   rects(dwData) = lprcMonitor
   Dim MI As MONITORINFO
   MI.cbSize = Len(MI)
   If apiGetMonitorInfo(hMonitor, MI) <> 0 Then
      With rects(dwData)
         Dim r As New Rectangle
         r.left = MI.rcMonitor.left
         r.top = MI.rcMonitor.top
         r.right = MI.rcMonitor.right
         r.bottom = MI.rcMonitor.bottom
         r.Width_ = (MI.rcMonitor.right - MI.rcMonitor.left)
         r.Height = (MI.rcMonitor.bottom - MI.rcMonitor.top)
         r.Size.Width_ = r.Width_
         r.Size.Height_ = r.Height
         r.Location.x = r.left
         r.Location.y = r.top
         Dim rw As New Rectangle
         With rw
            .left = MI.rcWork.left
            .top = MI.rcWork.top
            .right = MI.rcWork.right
            .bottom = MI.rcWork.bottom
            .Width_ = (MI.rcWork.right - MI.rcWork.left)
            .Height = (MI.rcWork.bottom - MI.rcWork.top)
            .Size.Height_ = rw.Width_
            .Size.Width_ = rw.Height
            .Location.x = rw.left
            .Location.y = rw.top
         End With
         Dim sc As New Screen
         If hMon = 0 Then sc.Primary = CBool(MI.dwFlags = MONITORINFOF_PRIMARY)
         If hMon <> 0 And hMon = hMonitor Then sc.Primary = True
         Set sc.Bounds = r
         Set sc.WorkingArea = rw
         Set screens(dwData) = sc
      End With
   End If
   dwData = dwData + 1
   MonitorEnumProc = 1
End Function
