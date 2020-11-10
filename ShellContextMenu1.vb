Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Windows.Forms
Imports System.IO
Imports System.Security.Permissions
Namespace ShellContext
    ''' Create an instance and call ShowContextMenu with a list of FileInfo for the files.

    ''' Hooking class taken from MSDN Magazine Cutting Edge column
    ''' http://msdn.microsoft.com/msdnmag/issues/02/10/CuttingEdge/
    Public Class ShellContextMenu
        Inherits NativeWindow
        Private Const MAX_PATH As Integer = 260
        Private Const CMD_FIRST As UInteger = 1
        Private Const CMD_LAST As UInteger = 30000
        Private Const S_OK As Integer = 0
        Private Const S_FALSE As Integer = 1
        Private Shared cbMenuItemInfo As Integer = Marshal.SizeOf(GetType(MENUITEMINFO))
        Private Shared cbInvokeCommand As Integer = Marshal.SizeOf(GetType(CMINVOKECOMMANDINFOEX))
        Private Shared IID_IShellFolder As New Guid("{000214E6-0000-0000-C000-000000000046}")
        Private Shared IID_IContextMenu As New Guid("{000214e4-0000-0000-c000-000000000046}")
        Private Shared IID_IContextMenu2 As New Guid("{000214f4-0000-0000-c000-000000000046}")
        Private Shared IID_IContextMenu3 As New Guid("{bcfce0a0-ec17-11d0-8d10-00a0c90f2719}")
        <StructLayout(LayoutKind.Sequential)> _
        Private Structure CWPSTRUCT
            Public lparam As IntPtr
            Public wparam As IntPtr
            Public message As Integer
            Public hwnd As IntPtr
        End Structure
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
        Private Structure CMINVOKECOMMANDINFOEX ' Contains extended information about a shortcut menu command
            Public cbSize As Integer
            Public fMask As CMIC
            Public hwnd As IntPtr
            Public lpVerb As IntPtr
            <MarshalAs(UnmanagedType.LPStr)> _
            Public lpParameters As String
            <MarshalAs(UnmanagedType.LPStr)> _
            Public lpDirectory As String
            Public nShow As SW
            Public dwHotKey As Integer
            Public hIcon As IntPtr
            <MarshalAs(UnmanagedType.LPStr)> _
            Public lpTitle As String
            Public lpVerbW As IntPtr
            <MarshalAs(UnmanagedType.LPWStr)> _
            Public lpParametersW As String
            <MarshalAs(UnmanagedType.LPWStr)> _
            Public lpDirectoryW As String
            <MarshalAs(UnmanagedType.LPWStr)> _
            Public lpTitleW As String
            Public ptInvoke As POINT
        End Structure
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)> _
        Private Structure MENUITEMINFO  ' Contains information about a menu item
            Public Sub New(ByVal text As String)
                cbSize = cbMenuItemInfo
                dwTypeData = text
                cch = text.Length
                fMask = 0
                fType = 0
                fState = 0
                wID = 0
                hSubMenu = IntPtr.Zero
                hbmpChecked = IntPtr.Zero
                hbmpUnchecked = IntPtr.Zero
                dwItemData = IntPtr.Zero
                hbmpItem = IntPtr.Zero
            End Sub
            Public cbSize As Integer
            Public fMask As MIIM
            Public fType As MFT
            Public fState As MFS
            Public wID As UInteger
            Public hSubMenu As IntPtr
            Public hbmpChecked As IntPtr
            Public hbmpUnchecked As IntPtr
            Public dwItemData As IntPtr
            <MarshalAs(UnmanagedType.LPTStr)> _
            Public dwTypeData As String
            Public cch As Integer
            Public hbmpItem As IntPtr
        End Structure
        <StructLayout(LayoutKind.Sequential)> _
        Private Structure STGMEDIUM ' A generalized global memory handle used for data transfer operations by the IAdviseSink, IDataObject, and IOleCache interfaces
            Public tymed As TYMED
            Public hBitmap As IntPtr
            Public hMetaFilePict As IntPtr
            Public hEnhMetaFile As IntPtr
            Public hGlobal As IntPtr
            Public lpszFileName As IntPtr
            Public pstm As IntPtr
            Public pstg As IntPtr
            Public pUnkForRelease As IntPtr
        End Structure
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)> _
        Private Structure POINT ' Defines the x- and y-coordinates of a point
            Public Sub New(ByVal x As Integer, ByVal y As Integer)
                Me.x = x
                Me.y = y
            End Sub
            Public x As Integer
            Public y As Integer
        End Structure
        <Flags()> _
        Private Enum SHGNO ' Defines the values used with the IShellFolder::GetDisplayNameOf and IShellFolder::SetNameOf  methods to specify the type of file or folder names used by those methods
            NORMAL = &H0
            INFOLDER = &H1
            FOREDITING = &H1000
            FORADDRESSBAR = &H4000
            FORPARSING = &H8000
        End Enum
        <Flags()> _
        Private Enum SFGAO As UInteger ' The attributes that the caller is requesting, when calling IShellFolder::GetAttributesOf
            BROWSABLE = &H8000000
            CANCOPY = 1
            CANDELETE = &H20
            CANLINK = 4
            CANMONIKER = &H400000
            CANMOVE = 2
            CANRENAME = &H10
            CAPABILITYMASK = &H177
            COMPRESSED = &H4000000
            CONTENTSMASK = &H80000000UI
            DISPLAYATTRMASK = &HFC000
            DROPTARGET = &H100
            ENCRYPTED = &H2000
            FILESYSANCESTOR = &H10000000
            FILESYSTEM = &H40000000
            FOLDER = &H20000000
            GHOSTED = &H8000
            HASPROPSHEET = &H40
            HASSTORAGE = &H400000
            HASSUBFOLDER = &H80000000UI
            HIDDEN = &H80000
            ISSLOW = &H4000
            LINK = &H10000
            NEWCONTENT = &H200000
            NONENUMERATED = &H100000
            [READONLY] = &H40000
            REMOVABLE = &H2000000
            SHARE = &H20000
            STORAGE = 8
            STORAGEANCESTOR = &H800000
            STORAGECAPMASK = &H70C50008
            STREAM = &H400000
            VALIDATE = &H1000000
        End Enum
        <Flags()> _
        Private Enum SHCONTF ' Determines the type of items included in an enumeration.  These values are used with the IShellFolder::EnumObjects method
            FOLDERS = &H20
            NONFOLDERS = &H40
            INCLUDEHIDDEN = &H80
            INIT_ON_FIRST_NEXT = &H100
            NETPRINTERSRCH = &H200
            SHAREABLE = &H400
            STORAGE = &H800
        End Enum
        <Flags()> _
        Private Enum CMF As UInteger ' Specifies how the shortcut menu can be changed when calling IContextMenu::QueryContextMenu
            NORMAL = &H0
            DEFAULTONLY = &H1
            VERBSONLY = &H2
            EXPLORE = &H4
            NOVERBS = &H8
            CANRENAME = &H10
            NODEFAULT = &H20
            INCLUDESTATIC = &H40
            EXTENDEDVERBS = &H100
            RESERVED = &HFFFF0000UI
        End Enum
        <Flags()> _
        Private Enum GCS As UInteger  ' Flags specifying the information to return when calling IContextMenu::GetCommandString
            VERBA = 0
            HELPTEXTA = 1
            VALIDATEA = 2
            VERBW = 4
            HELPTEXTW = 5
            VALIDATEW = 6
        End Enum
        <Flags()> _
        Private Enum TPM As UInteger        ' Specifies how TrackPopupMenuEx positions the shortcut menu horizontally
            LEFTBUTTON = &H0
            RIGHTBUTTON = &H2
            LEFTALIGN = &H0
            CENTERALIGN = &H4
            RIGHTALIGN = &H8
            TOPALIGN = &H0
            VCENTERALIGN = &H10
            BOTTOMALIGN = &H20
            HORIZONTAL = &H0
            VERTICAL = &H40
            NONOTIFY = &H80
            RETURNCMD = &H100
            RECURSE = &H1
            HORPOSANIMATION = &H400
            HORNEGANIMATION = &H800
            VERPOSANIMATION = &H1000
            VERNEGANIMATION = &H2000
            NOANIMATION = &H4000
            LAYOUTRTL = &H8000
        End Enum
        Private Enum CMD_CUSTOM ' The cmd for a custom added menu item
            ExpandCollapse = CInt(CMD_LAST) + 1
        End Enum
        <Flags()> _
        Private Enum CMIC As UInteger ' Flags used with the CMINVOKECOMMANDINFOEX structure
            HOTKEY = &H20
            ICON = &H10
            FLAG_NO_UI = &H400
            UNICODE = &H4000
            NO_CONSOLE = &H8000
            ASYNCOK = &H100000
            NOZONECHECKS = &H800000
            SHIFT_DOWN = &H10000000
            CONTROL_DOWN = &H40000000
            FLAG_LOG_USAGE = &H4000000
            PTINVOKE = &H20000000
        End Enum
        <Flags()> _
        Private Enum SW  ' Specifies how the window is to be shown
            HIDE = 0
            SHOWNORMAL = 1
            NORMAL = 1
            SHOWMINIMIZED = 2
            SHOWMAXIMIZED = 3
            MAXIMIZE = 3
            SHOWNOACTIVATE = 4
            SHOW = 5
            MINIMIZE = 6
            SHOWMINNOACTIVE = 7
            SHOWNA = 8
            RESTORE = 9
            SHOWDEFAULT = 10
        End Enum
        <Flags()> _
        Private Enum WM As UInteger 'window messages
            ACTIVATE = &H6
            ACTIVATEAPP = &H1C
            AFXFIRST = &H360
            AFXLAST = &H37F
            APP = &H8000
            ASKCBFORMATNAME = &H30C
            CANCELJOURNAL = &H4B
            CANCELMODE = &H1F
            CAPTURECHANGED = &H215
            CHANGECBCHAIN = &H30D
            [CHAR] = &H102
            CHARTOITEM = &H2F
            CHILDACTIVATE = &H22
            CLEAR = &H303
            CLOSE = &H10
            COMMAND = &H111
            COMPACTING = &H41
            COMPAREITEM = &H39
            CONTEXTMENU = &H7B
            COPY = &H301
            COPYDATA = &H4A
            CREATE = &H1
            CTLCOLORBTN = &H135
            CTLCOLORDLG = &H136
            CTLCOLOREDIT = &H133
            CTLCOLORLISTBOX = &H134
            CTLCOLORMSGBOX = &H132
            CTLCOLORSCROLLBAR = &H137
            CTLCOLORSTATIC = &H138
            CUT = &H300
            DEADCHAR = &H103
            DELETEITEM = &H2D
            DESTROY = &H2
            DESTROYCLIPBOARD = &H307
            DEVICECHANGE = &H219
            DEVMODECHANGE = &H1B
            DISPLAYCHANGE = &H7E
            DRAWCLIPBOARD = &H308
            DRAWITEM = &H2B
            DROPFILES = &H233
            ENABLE = &HA
            ENDSESSION = &H16
            ENTERIDLE = &H121
            ENTERMENULOOP = &H211
            ENTERSIZEMOVE = &H231
            ERASEBKGND = &H14
            EXITMENULOOP = &H212
            EXITSIZEMOVE = &H232
            FONTCHANGE = &H1D
            GETDLGCODE = &H87
            GETFONT = &H31
            GETHOTKEY = &H33
            GETICON = &H7F
            GETMINMAXINFO = &H24
            GETOBJECT = &H3D
            GETSYSMENU = &H313
            GETTEXT = &HD
            GETTEXTLENGTH = &HE
            HANDHELDFIRST = &H358
            HANDHELDLAST = &H35F
            HELP = &H53
            HOTKEY = &H312
            HSCROLL = &H114
            HSCROLLCLIPBOARD = &H30E
            ICONERASEBKGND = &H27
            IME_CHAR = &H286
            IME_COMPOSITION = &H10F
            IME_COMPOSITIONFULL = &H284
            IME_CONTROL = &H283
            IME_ENDCOMPOSITION = &H10E
            IME_KEYDOWN = &H290
            IME_KEYLAST = &H10F
            IME_KEYUP = &H291
            IME_NOTIFY = &H282
            IME_REQUEST = &H288
            IME_SELECT = &H285
            IME_SETCONTEXT = &H281
            IME_STARTCOMPOSITION = &H10D
            INITDIALOG = &H110
            INITMENU = &H116
            INITMENUPOPUP = &H117
            INPUTLANGCHANGE = &H51
            INPUTLANGCHANGEREQUEST = &H50
            KEYDOWN = &H100
            KEYFIRST = &H100
            KEYLAST = &H108
            KEYUP = &H101
            KILLFOCUS = &H8
            LBUTTONDBLCLK = &H203
            LBUTTONDOWN = &H201
            LBUTTONUP = &H202
            LVM_GETEDITCONTROL = &H1018
            LVM_SETIMAGELIST = &H1003
            MBUTTONDBLCLK = &H209
            MBUTTONDOWN = &H207
            MBUTTONUP = &H208
            MDIACTIVATE = &H222
            MDICASCADE = &H227
            MDICREATE = &H220
            MDIDESTROY = &H221
            MDIGETACTIVE = &H229
            MDIICONARRANGE = &H228
            MDIMAXIMIZE = &H225
            MDINEXT = &H224
            MDIREFRESHMENU = &H234
            MDIRESTORE = &H223
            MDISETMENU = &H230
            MDITILE = &H226
            MEASUREITEM = &H2C
            MENUCHAR = &H120
            MENUCOMMAND = &H126
            MENUDRAG = &H123
            MENUGETOBJECT = &H124
            MENURBUTTONUP = &H122
            MENUSELECT = &H11F
            MOUSEACTIVATE = &H21
            MOUSEFIRST = &H200
            MOUSEHOVER = &H2A1
            MOUSELAST = &H20A
            MOUSELEAVE = &H2A3
            MOUSEMOVE = &H200
            MOUSEWHEEL = &H20A
            MOVE = &H3
            MOVING = &H216
            NCACTIVATE = &H86
            NCCALCSIZE = &H83
            NCCREATE = &H81
            NCDESTROY = &H82
            NCHITTEST = &H84
            NCLBUTTONDBLCLK = &HA3
            NCLBUTTONDOWN = &HA1
            NCLBUTTONUP = &HA2
            NCMBUTTONDBLCLK = &HA9
            NCMBUTTONDOWN = &HA7
            NCMBUTTONUP = &HA8
            NCMOUSEHOVER = &H2A0
            NCMOUSELEAVE = &H2A2
            NCMOUSEMOVE = &HA0
            NCPAINT = &H85
            NCRBUTTONDBLCLK = &HA6
            NCRBUTTONDOWN = &HA4
            NCRBUTTONUP = &HA5
            NEXTDLGCTL = &H28
            NEXTMENU = &H213
            NOTIFY = &H4E
            NOTIFYFORMAT = &H55
            NULL = &H0
            PAINT = &HF
            PAINTCLIPBOARD = &H309
            PAINTICON = &H26
            PALETTECHANGED = &H311
            PALETTEISCHANGING = &H310
            PARENTNOTIFY = &H210
            PASTE = &H302
            PENWINFIRST = &H380
            PENWINLAST = &H38F
            POWER = &H48
            PRINT = &H317
            PRINTCLIENT = &H318
            QUERYDRAGICON = &H37
            QUERYENDSESSION = &H11
            QUERYNEWPALETTE = &H30F
            QUERYOPEN = &H13
            QUEUESYNC = &H23
            QUIT = &H12
            RBUTTONDBLCLK = &H206
            RBUTTONDOWN = &H204
            RBUTTONUP = &H205
            RENDERALLFORMATS = &H306
            RENDERFORMAT = &H305
            SETCURSOR = &H20
            SETFOCUS = &H7
            SETFONT = &H30
            SETHOTKEY = &H32
            SETICON = &H80
            SETMARGINS = &HD3
            SETREDRAW = &HB
            SETTEXT = &HC
            SETTINGCHANGE = &H1A
            SHOWWINDOW = &H18
            SIZE = &H5
            SIZECLIPBOARD = &H30B
            SIZING = &H214
            SPOOLERSTATUS = &H2A
            STYLECHANGED = &H7D
            STYLECHANGING = &H7C
            SYNCPAINT = &H88
            SYSCHAR = &H106
            SYSCOLORCHANGE = &H15
            SYSCOMMAND = &H112
            SYSDEADCHAR = &H107
            SYSKEYDOWN = &H104
            SYSKEYUP = &H105
            TCARD = &H52
            TIMECHANGE = &H1E
            TIMER = &H113
            TVM_GETEDITCONTROL = &H110F
            TVM_SETIMAGELIST = &H1109
            UNDO = &H304
            UNINITMENUPOPUP = &H125
            USER = &H400
            USERCHANGED = &H54
            VKEYTOITEM = &H2E
            VSCROLL = &H115
            VSCROLLCLIPBOARD = &H30A
            WINDOWPOSCHANGED = &H47
            WINDOWPOSCHANGING = &H46
            WININICHANGE = &H1A
            SH_NOTIFY = &H401
        End Enum
        <Flags()> _
        Private Enum MFT As UInteger    ' Specifies the content of the new menu item
            GRAYED = &H3
            DISABLED = &H3
            CHECKED = &H8
            SEPARATOR = &H800
            RADIOCHECK = &H200
            BITMAP = &H4
            OWNERDRAW = &H100
            MENUBARBREAK = &H20
            MENUBREAK = &H40
            RIGHTORDER = &H2000
            BYCOMMAND = &H0
            BYPOSITION = &H400
            POPUP = &H10
        End Enum
        <Flags()> _
        Private Enum MFS As UInteger    ' Specifies the state of the new menu item
            GRAYED = &H3
            DISABLED = &H3
            CHECKED = &H8
            HILITE = &H80
            ENABLED = &H0
            UNCHECKED = &H0
            UNHILITE = &H0
            [DEFAULT] = &H1000
        End Enum
        <Flags()> _
        Private Enum MIIM As UInteger ' Specifies the content of the new menu item
            BITMAP = &H80
            CHECKMARKS = &H8
            DATA = &H20
            FTYPE = &H100
            ID = &H2
            STATE = &H1
            [STRING] = &H40
            SUBMENU = &H4
            TYPE = &H10
        End Enum
        <Flags()> _
        Private Enum TYMED ' Indicates the type of storage medium being used in a data transfer
            ENHMF = &H40
            FILE = 2
            GDI = &H10
            HGLOBAL = 1
            ISTORAGE = 8
            ISTREAM = 4
            MFPICT = &H20
            NULL = 0
        End Enum
        <DllImport("shell32.dll")> _
        Private Shared Function SHGetDesktopFolder(ByRef ppshf As IntPtr) As Int32 ' Retrieves the IShellFolder interface for the desktop folder, which is the root of the Shell's namespace.
        End Function
        <DllImport("shlwapi.dll", EntryPoint:="StrRetToBuf", ExactSpelling:=False, CharSet:=CharSet.Auto, SetLastError:=True)> _
        Private Shared Function StrRetToBuf(ByVal pstr As IntPtr, ByVal pidl As IntPtr, ByVal pszBuf As StringBuilder, ByVal cchBuf As Integer) As Int32 ' Takes a STRRET structure returned by IShellFolder::GetDisplayNameOf, converts it to a string, and places the result in a buffer. 
        End Function
        <DllImport("user32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto)> _
        Private Shared Function TrackPopupMenuEx(ByVal hmenu As IntPtr, ByVal flags As TPM, ByVal x As Integer, ByVal y As Integer, ByVal hwnd As IntPtr, ByVal lptpm As IntPtr) As UInteger ' The TrackPopupMenuEx function displays a shortcut menu at the specified location and tracks the selection of items on the shortcut menu. The shortcut menu can appear anywhere on the screen.
        End Function
        <DllImport("user32", SetLastError:=True, CharSet:=CharSet.Auto)> _
        Private Shared Function CreatePopupMenu() As IntPtr  ' The CreatePopupMenu function creates a drop-down menu, submenu, or shortcut menu. The menu is initially empty. You can insert or append menu items by using the InsertMenuItem function. You can also use the InsertMenu function to insert menu items and the AppendMenu function to append menu items.
        End Function
        <DllImport("user32", SetLastError:=True, CharSet:=CharSet.Auto)> _
        Private Shared Function DestroyMenu(ByVal hMenu As IntPtr) As Boolean   ' The DestroyMenu function destroys the specified menu and frees any memory that the menu occupies.
        End Function
        <DllImport("user32", SetLastError:=True, CharSet:=CharSet.Auto)> _
        Private Shared Function GetMenuDefaultItem(ByVal hMenu As IntPtr, ByVal fByPos As Boolean, ByVal gmdiFlags As UInteger) As Integer   ' Determines the default menu item on the specified menu
        End Function
        Private _oContextMenu As IContextMenu
        Private _oContextMenu2 As IContextMenu2
        Private _oContextMenu3 As IContextMenu3
        Private _oDesktopFolder As IShellFolder
        Private _oParentFolder As IShellFolder
        Private _arrPIDLs() As IntPtr
        Private _strParentFolder As String
        Public Sub New()
            Me.CreateHandle(New CreateParams())
        End Sub
        Protected Overrides Sub Finalize()
            ReleaseAll()
        End Sub

        Private Sub InvokeContextMenuDefault(ByVal arrFI() As FileInfo)
            ReleaseAll()  ' Release all resources first.
            Dim pMenu As IntPtr = IntPtr.Zero, iContextMenuPtr As IntPtr = IntPtr.Zero

            Try
                _arrPIDLs = GetPIDLs(arrFI)
                If Nothing Is _arrPIDLs Then
                    ReleaseAll()
                    Return
                End If

                If False = GetContextMenuInterfaces(_oParentFolder, _arrPIDLs, iContextMenuPtr) Then
                    ReleaseAll()
                    Return
                End If

                pMenu = CreatePopupMenu()

                Dim nResult As Integer = _oContextMenu.QueryContextMenu(pMenu, 0, CMD_FIRST, CMD_LAST, CMF.DEFAULTONLY) 'Or (If((Control.ModifierKeys And Keys.Shift) <> 0, CMF.EXTENDEDVERBS, 0)))

                Dim nDefaultCmd As UInteger = CUInt(Math.Truncate(GetMenuDefaultItem(pMenu, False, 0)))
                If nDefaultCmd >= CMD_FIRST Then
                    InvokeCommand(_oContextMenu, nDefaultCmd, arrFI(0).DirectoryName, Control.MousePosition)


                End If

                DestroyMenu(pMenu)
                pMenu = IntPtr.Zero
            Catch
                Throw
            Finally
                If pMenu <> IntPtr.Zero Then
                    DestroyMenu(pMenu)
                End If
                ReleaseAll()
            End Try
        End Sub
        Private Sub InvokeCommand(ByVal oContextMenu As IContextMenu, ByVal nCmd As UInteger, ByVal strFolder As String, ByVal pointInvoke As Drawing.Point)
            On Error Resume Next
            Dim invoke As New CMINVOKECOMMANDINFOEX()
            invoke.cbSize = cbInvokeCommand
            invoke.lpVerb = CType(CInt(nCmd) - CMD_FIRST, IntPtr)
            invoke.lpDirectory = strFolder
            invoke.lpVerbW = CType(CInt(nCmd) - CMD_FIRST, IntPtr)
            invoke.lpDirectoryW = strFolder
            invoke.fMask = CMIC.UNICODE Or CMIC.PTINVOKE 'Or (If((Control.ModifierKeys And Keys.Control) <> 0, CMIC.CONTROL_DOWN, 0)) Or (If((Control.ModifierKeys And Keys.Shift) <> 0, CMIC.SHIFT_DOWN, 0))
            invoke.ptInvoke = New POINT(pointInvoke.X, pointInvoke.Y)
            invoke.nShow = SW.SHOWNORMAL
            oContextMenu.InvokeCommand(invoke)

        End Sub
        ''' <summary>Gets the interfaces to the context menu</summary>
        ''' <param name="oParentFolder">Parent folder</param>
        ''' <param name="arrPIDLs">PIDLs</param>
        ''' <returns>true if it got the interfaces, otherwise false</returns>
        Private Function GetContextMenuInterfaces(ByVal oParentFolder As IShellFolder, ByVal arrPIDLs() As IntPtr, ByRef ctxMenuPtr As IntPtr) As Boolean
            Dim nResult As Integer = oParentFolder.GetUIObjectOf(IntPtr.Zero, CUInt(arrPIDLs.Length), arrPIDLs, IID_IContextMenu, IntPtr.Zero, ctxMenuPtr)
            If S_OK = nResult Then
                _oContextMenu = DirectCast(Marshal.GetTypedObjectForIUnknown(ctxMenuPtr, GetType(IContextMenu)), IContextMenu)
                Return True
            Else
                ctxMenuPtr = IntPtr.Zero
                _oContextMenu = Nothing
                Return False
            End If
        End Function
        ''' <summary>
        ''' This method receives WindowMessages. It will make the "Open With" and "Send To" work 
        ''' by calling HandleMenuMsg and HandleMenuMsg2. It will also call the OnContextMenuMouseHover 
        ''' method of Browser when hovering over a ContextMenu item.
        ''' </summary>
        ''' <param name="m">the Message of the Browser's WndProc</param>
        ''' <returns>true if the message has been handled, false otherwise</returns>
        Protected Overrides Sub WndProc(ByRef m As Message)
            '			#Region "IContextMenu"

            If _oContextMenu IsNot Nothing AndAlso m.Msg = CInt(WM.MENUSELECT) AndAlso (CInt(ShellHelper.HiWord(m.WParam)) And CInt(MFT.SEPARATOR)) = 0 AndAlso (CInt(ShellHelper.HiWord(m.WParam)) And CInt(MFT.POPUP)) = 0 Then
                Dim info As String = String.Empty

                If ShellHelper.LoWord(m.WParam) = CInt(CMD_CUSTOM.ExpandCollapse) Then
                    info = "Expands or collapses the current selected item"
                Else
                    info = ""
                End If
            End If

            '			#End Region

            '			#Region "IContextMenu2"

            If _oContextMenu2 IsNot Nothing AndAlso (m.Msg = CInt(WM.INITMENUPOPUP) OrElse m.Msg = CInt(WM.MEASUREITEM) OrElse m.Msg = CInt(WM.DRAWITEM)) Then
                If _oContextMenu2.HandleMenuMsg(CUInt(m.Msg), m.WParam, m.LParam) = S_OK Then
                    Return
                End If
            End If

            '			#End Region

            '			#Region "IContextMenu3"

            If _oContextMenu3 IsNot Nothing AndAlso m.Msg = CInt(WM.MENUCHAR) Then
                If _oContextMenu3.HandleMenuMsg2(CUInt(m.Msg), m.WParam, m.LParam, IntPtr.Zero) = S_OK Then
                    Return
                End If
            End If

            '			#End Region

            MyBase.WndProc(m)
        End Sub
        ''' <summary>
        ''' Release all allocated interfaces, PIDLs 
        ''' </summary>
        Friend Sub ReleaseAll()
            If Nothing IsNot _oContextMenu Then
                Marshal.ReleaseComObject(_oContextMenu)
                _oContextMenu = Nothing
            End If
            If Nothing IsNot _oContextMenu2 Then
                Marshal.ReleaseComObject(_oContextMenu2)
                _oContextMenu2 = Nothing
            End If
            If Nothing IsNot _oContextMenu3 Then
                Marshal.ReleaseComObject(_oContextMenu3)
                _oContextMenu3 = Nothing
            End If
            If Nothing IsNot _oDesktopFolder Then
                Marshal.ReleaseComObject(_oDesktopFolder)
                _oDesktopFolder = Nothing
            End If
            If Nothing IsNot _oParentFolder Then
                Marshal.ReleaseComObject(_oParentFolder)
                _oParentFolder = Nothing
            End If
            If Nothing IsNot _arrPIDLs Then
                FreePIDLs(_arrPIDLs)
                _arrPIDLs = Nothing
            End If
        End Sub
        ''' <summary>
        ''' Gets the desktop folder
        ''' </summary>
        ''' <returns>IShellFolder for desktop folder</returns>
        Private Function GetDesktopFolder() As IShellFolder
            Dim pUnkownDesktopFolder As IntPtr = IntPtr.Zero

            If Nothing Is _oDesktopFolder Then
                ' Get desktop IShellFolder
                Dim nResult As Integer = SHGetDesktopFolder(pUnkownDesktopFolder)
                If S_OK <> nResult Then
                    Throw New ShellContextMenuException("Failed to get the desktop shell folder")
                End If
                _oDesktopFolder = DirectCast(Marshal.GetTypedObjectForIUnknown(pUnkownDesktopFolder, GetType(IShellFolder)), IShellFolder)
            End If

            Return _oDesktopFolder
        End Function
        ''' <summary>
        ''' Gets the parent folder
        ''' </summary>
        ''' <param name="folderName">Folder path</param>
        ''' <returns>IShellFolder for the folder (relative from the desktop)</returns>
        Private Function GetParentFolder(ByVal folderName As String) As IShellFolder
            If Nothing Is _oParentFolder Then
                Dim oDesktopFolder As IShellFolder = GetDesktopFolder()
                If Nothing Is oDesktopFolder Then
                    Return Nothing
                End If
                Dim pPIDL As IntPtr = IntPtr.Zero ' Get the PIDL for the folder file is in
                Dim pchEaten As UInteger = 0
                Dim pdwAttributes As SFGAO = 0
                Dim nResult As Integer = oDesktopFolder.ParseDisplayName(IntPtr.Zero, IntPtr.Zero, folderName, pchEaten, pPIDL, pdwAttributes)
                If S_OK <> nResult Then
                    Return Nothing
                End If
                Dim pStrRet As IntPtr = Marshal.AllocCoTaskMem(MAX_PATH * 2 + 4)
                Marshal.WriteInt32(pStrRet, 0, 0)
                nResult = _oDesktopFolder.GetDisplayNameOf(pPIDL, SHGNO.FORPARSING, pStrRet)
                Dim strFolder As New StringBuilder(MAX_PATH)
                StrRetToBuf(pStrRet, pPIDL, strFolder, MAX_PATH)
                Marshal.FreeCoTaskMem(pStrRet)
                pStrRet = IntPtr.Zero
                _strParentFolder = strFolder.ToString()
                Dim pUnknownParentFolder As IntPtr = IntPtr.Zero ' Get the IShellFolder for folder
                nResult = oDesktopFolder.BindToObject(pPIDL, IntPtr.Zero, IID_IShellFolder, pUnknownParentFolder)
                Marshal.FreeCoTaskMem(pPIDL) ' Free the PIDL first
                If S_OK <> nResult Then
                    Return Nothing
                End If
                _oParentFolder = DirectCast(Marshal.GetTypedObjectForIUnknown(pUnknownParentFolder, GetType(IShellFolder)), IShellFolder)
            End If
            Return _oParentFolder
        End Function
        ''' <summary>
        ''' Get the PIDLs
        ''' </summary>
        ''' <param name="arrFI">Array of FileInfo</param>
        ''' <returns>Array of PIDLs</returns>
        Protected Function GetPIDLs(ByVal arrFI() As FileInfo) As IntPtr()
            If Nothing Is arrFI OrElse 0 = arrFI.Length Then
                Return Nothing
            End If

            Dim oParentFolder As IShellFolder = GetParentFolder(arrFI(0).DirectoryName)
            If Nothing Is oParentFolder Then
                Return Nothing
            End If

            Dim arrPIDLs(arrFI.Length - 1) As IntPtr
            Dim n As Integer = 0
            For Each fi As FileInfo In arrFI
                ' Get the file relative to folder
                Dim pchEaten As UInteger = 0
                Dim pdwAttributes As SFGAO = 0
                Dim pPIDL As IntPtr = IntPtr.Zero
                Dim nResult As Integer = oParentFolder.ParseDisplayName(IntPtr.Zero, IntPtr.Zero, fi.Name, pchEaten, pPIDL, pdwAttributes)
                If S_OK <> nResult Then
                    FreePIDLs(arrPIDLs)
                    Return Nothing
                End If
                arrPIDLs(n) = pPIDL
                n += 1
            Next fi

            Return arrPIDLs
        End Function
        ''' <summary>
        ''' Get the PIDLs
        ''' </summary>
        ''' <param name="arrFI">Array of DirectoryInfo</param>
        ''' <returns>Array of PIDLs</returns>
        Protected Function GetPIDLs(ByVal arrFI() As DirectoryInfo) As IntPtr()
            If Nothing Is arrFI OrElse 0 = arrFI.Length Then
                Return Nothing
            End If

            Dim oParentFolder As IShellFolder = GetParentFolder(arrFI(0).Parent.FullName)
            If Nothing Is oParentFolder Then
                Return Nothing
            End If

            Dim arrPIDLs(arrFI.Length - 1) As IntPtr
            Dim n As Integer = 0
            For Each fi As DirectoryInfo In arrFI
                ' Get the file relative to folder
                Dim pchEaten As UInteger = 0
                Dim pdwAttributes As SFGAO = 0
                Dim pPIDL As IntPtr = IntPtr.Zero
                Dim nResult As Integer = oParentFolder.ParseDisplayName(IntPtr.Zero, IntPtr.Zero, fi.Name, pchEaten, pPIDL, pdwAttributes)
                If S_OK <> nResult Then
                    FreePIDLs(arrPIDLs)
                    Return Nothing
                End If
                arrPIDLs(n) = pPIDL
                n += 1
            Next fi

            Return arrPIDLs
        End Function
        ''' <summary>
        ''' Free the PIDLs
        ''' </summary>
        ''' <param name="arrPIDLs">Array of PIDLs (IntPtr)</param>
        Protected Sub FreePIDLs(ByVal arrPIDLs() As IntPtr)
            If Nothing IsNot arrPIDLs Then
                For n As Integer = 0 To arrPIDLs.Length - 1
                    If arrPIDLs(n) <> IntPtr.Zero Then
                        Marshal.FreeCoTaskMem(arrPIDLs(n))
                        arrPIDLs(n) = IntPtr.Zero
                    End If
                Next n
            End If
        End Sub
        ''' <summary>
        ''' Shows the context menu
        ''' </summary>
        ''' <param name="files">FileInfos (should all be in same directory)</param>
        ''' <param name="pointScreen">Where to show the menu</param>
        Public Sub ShowContextMenu(ByVal files() As FileInfo, ByVal pointScreen As Drawing.Point)
            ' Release all resources first.
            ReleaseAll()
            _arrPIDLs = GetPIDLs(files)
            Me.ShowContextMenu(pointScreen)
        End Sub
        ''' <summary>
        ''' Shows the context menu
        ''' </summary>
        ''' <param name="dirs">DirectoryInfos (should all be in same directory)</param>
        ''' <param name="pointScreen">Where to show the menu</param>
        Public Sub ShowContextMenu(ByVal dirs() As DirectoryInfo, ByVal pointScreen As Drawing.Point)
            ' Release all resources first.
            ReleaseAll()
            _arrPIDLs = GetPIDLs(dirs)
            Me.ShowContextMenu(pointScreen)
        End Sub
        ''' <summary>
        ''' Shows the context menu
        ''' </summary>
        ''' <param name="arrFI">FileInfos (should all be in same directory)</param>
        ''' <param name="pointScreen">Where to show the menu</param>
        Private Sub ShowContextMenu(ByVal pointScreen As Drawing.Point)
            Dim pMenu As IntPtr = IntPtr.Zero, iContextMenuPtr As IntPtr = IntPtr.Zero, iContextMenuPtr2 As IntPtr = IntPtr.Zero, iContextMenuPtr3 As IntPtr = IntPtr.Zero

            Try
                If Nothing Is _arrPIDLs Then
                    ReleaseAll()
                    Return
                End If

                If False = GetContextMenuInterfaces(_oParentFolder, _arrPIDLs, iContextMenuPtr) Then
                    ReleaseAll()
                    Return
                End If

                pMenu = CreatePopupMenu()

                Dim nResult As Integer = _oContextMenu.QueryContextMenu(pMenu, 0, CMD_FIRST, CMD_LAST, CMF.EXPLORE Or CMF.NORMAL) 'Or (If((Control.ModifierKeys And Keys.Shift) <> 0, CMF.EXTENDEDVERBS, 0)))

                Marshal.QueryInterface(iContextMenuPtr, IID_IContextMenu2, iContextMenuPtr2)
                Marshal.QueryInterface(iContextMenuPtr, IID_IContextMenu3, iContextMenuPtr3)

                _oContextMenu2 = DirectCast(Marshal.GetTypedObjectForIUnknown(iContextMenuPtr2, GetType(IContextMenu2)), IContextMenu2)
                _oContextMenu3 = DirectCast(Marshal.GetTypedObjectForIUnknown(iContextMenuPtr3, GetType(IContextMenu3)), IContextMenu3)

                Dim nSelected As UInteger = TrackPopupMenuEx(pMenu, TPM.RETURNCMD, pointScreen.X, pointScreen.Y, Me.Handle, IntPtr.Zero)

                DestroyMenu(pMenu)
                pMenu = IntPtr.Zero

                If nSelected <> 0 Then
                    InvokeCommand(_oContextMenu, nSelected, _strParentFolder, pointScreen)
                End If
            Catch
                Throw
            Finally
                'hook.Uninstall();
                If pMenu <> IntPtr.Zero Then
                    DestroyMenu(pMenu)
                End If

                If iContextMenuPtr <> IntPtr.Zero Then
                    Marshal.Release(iContextMenuPtr)
                End If

                If iContextMenuPtr2 <> IntPtr.Zero Then
                    Marshal.Release(iContextMenuPtr2)
                End If

                If iContextMenuPtr3 <> IntPtr.Zero Then
                    Marshal.Release(iContextMenuPtr3)
                End If

                ReleaseAll()
            End Try
        End Sub
        <ComImport()> _
            <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
            <Guid("000214E6-0000-0000-C000-000000000046")> _
                  Private Interface IShellFolder
            ' Translates a file object's or folder's display name into an item identifier list.
            ' Return value: error code, if any
            <PreserveSig()> _
                 Function ParseDisplayName(ByVal hwnd As IntPtr, ByVal pbc As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> ByVal pszDisplayName As String, ByRef pchEaten As UInteger, ByRef ppidl As IntPtr, ByRef pdwAttributes As SFGAO) As Int32

            ' Allows a client to determine the contents of a folder by creating an item
            ' identifier enumeration object and returning its IEnumIDList interface.
            ' Return value: error code, if any
            <PreserveSig()> _
            Function EnumObjects(ByVal hwnd As IntPtr, ByVal grfFlags As SHCONTF, ByRef enumIDList As IntPtr) As Int32

            ' Retrieves an IShellFolder object for a subfolder.
            ' Return value: error code, if any
            <PreserveSig()> _
            Function BindToObject(ByVal pidl As IntPtr, ByVal pbc As IntPtr, ByRef riid As Guid, ByRef ppv As IntPtr) As Int32

            ' Requests a pointer to an object's storage interface. 
            ' Return value: error code, if any
            <PreserveSig()> _
            Function BindToStorage(ByVal pidl As IntPtr, ByVal pbc As IntPtr, ByRef riid As Guid, ByRef ppv As IntPtr) As Int32

            ' Determines the relative order of two file objects or folders, given their
            ' item identifier lists. Return value: If this method is successful, the
            ' CODE field of the HRESULT contains one of the following values (the code
            ' can be retrived using the helper function GetHResultCode): Negative A
            ' negative return value indicates that the first item should precede
            ' the second (pidl1 < pidl2). 

            ' Positive A positive return value indicates that the first item should
            ' follow the second (pidl1 > pidl2).  Zero A return value of zero
            ' indicates that the two items are the same (pidl1 = pidl2). 
            <PreserveSig()> _
            Function CompareIDs(ByVal lParam As IntPtr, ByVal pidl1 As IntPtr, ByVal pidl2 As IntPtr) As Int32

            ' Requests an object that can be used to obtain information from or interact
            ' with a folder object.
            ' Return value: error code, if any
            <PreserveSig()> _
            Function CreateViewObject(ByVal hwndOwner As IntPtr, ByVal riid As Guid, ByRef ppv As IntPtr) As Int32

            ' Retrieves the attributes of one or more file objects or subfolders. 
            ' Return value: error code, if any
            <PreserveSig()> _
            Function GetAttributesOf(ByVal cidl As UInteger, <MarshalAs(UnmanagedType.LPArray)> ByVal apidl() As IntPtr, ByRef rgfInOut As SFGAO) As Int32

            ' Retrieves an OLE interface that can be used to carry out actions on the
            ' specified file objects or folders.
            ' Return value: error code, if any
            <PreserveSig()> _
            Function GetUIObjectOf(ByVal hwndOwner As IntPtr, ByVal cidl As UInteger, <MarshalAs(UnmanagedType.LPArray)> ByVal apidl() As IntPtr, ByRef riid As Guid, ByVal rgfReserved As IntPtr, ByRef ppv As IntPtr) As Int32

            ' Retrieves the display name for the specified file object or subfolder. 
            ' Return value: error code, if any
            <PreserveSig()> _
            Function GetDisplayNameOf(ByVal pidl As IntPtr, ByVal uFlags As SHGNO, ByVal lpName As IntPtr) As Int32

            ' Sets the display name of a file object or subfolder, changing the item
            ' identifier in the process.
            ' Return value: error code, if any
            <PreserveSig()> _
            Function SetNameOf(ByVal hwnd As IntPtr, ByVal pidl As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> ByVal pszName As String, ByVal uFlags As SHGNO, ByRef ppidlOut As IntPtr) As Int32
        End Interface
        <ComImport()> _
        <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
        <GuidAttribute("000214e4-0000-0000-c000-000000000046")> _
        Private Interface IContextMenu
            ' Adds commands to a shortcut menu
            <PreserveSig()> _
            Function QueryContextMenu(ByVal hmenu As IntPtr, ByVal iMenu As UInteger, ByVal idCmdFirst As UInteger, ByVal idCmdLast As UInteger, ByVal uFlags As CMF) As Int32

            ' Carries out the command associated with a shortcut menu item
            <PreserveSig()> _
            Function InvokeCommand(ByRef info As CMINVOKECOMMANDINFOEX) As Int32

            ' Retrieves information about a shortcut menu command, 
            ' including the help string and the language-independent, 
            ' or canonical, name for the command
            <PreserveSig()> _
            Function GetCommandString(ByVal idcmd As UInteger, ByVal uflags As GCS, ByVal reserved As UInteger, <MarshalAs(UnmanagedType.LPArray)> ByVal commandstring() As Byte, ByVal cch As Integer) As Int32
        End Interface
        <ComImport(), Guid("000214f4-0000-0000-c000-000000000046")> _
        <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
        Private Interface IContextMenu2
            ' Adds commands to a shortcut menu
            <PreserveSig()> _
            Function QueryContextMenu(ByVal hmenu As IntPtr, ByVal iMenu As UInteger, ByVal idCmdFirst As UInteger, ByVal idCmdLast As UInteger, ByVal uFlags As CMF) As Int32

            ' Carries out the command associated with a shortcut menu item
            <PreserveSig()> _
            Function InvokeCommand(ByRef info As CMINVOKECOMMANDINFOEX) As Int32

            ' Retrieves information about a shortcut menu command, 
            ' including the help string and the language-independent, 
            ' or canonical, name for the command
            <PreserveSig()> _
            Function GetCommandString(ByVal idcmd As UInteger, ByVal uflags As GCS, ByVal reserved As UInteger, <MarshalAs(UnmanagedType.LPWStr)> ByVal commandstring As StringBuilder, ByVal cch As Integer) As Int32

            ' Allows client objects of the IContextMenu interface to 
            ' handle messages associated with owner-drawn menu items
            <PreserveSig()> _
            Function HandleMenuMsg(ByVal uMsg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Int32
        End Interface
        <ComImport(), Guid("bcfce0a0-ec17-11d0-8d10-00a0c90f2719")> _
        <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
        Private Interface IContextMenu3
            ' Adds commands to a shortcut menu
            <PreserveSig()> _
            Function QueryContextMenu(ByVal hmenu As IntPtr, ByVal iMenu As UInteger, ByVal idCmdFirst As UInteger, ByVal idCmdLast As UInteger, ByVal uFlags As CMF) As Int32

            ' Carries out the command associated with a shortcut menu item
            <PreserveSig()> _
            Function InvokeCommand(ByRef info As CMINVOKECOMMANDINFOEX) As Int32

            ' Retrieves information about a shortcut menu command, 
            ' including the help string and the language-independent, 
            ' or canonical, name for the command
            <PreserveSig()> _
            Function GetCommandString(ByVal idcmd As UInteger, ByVal uflags As GCS, ByVal reserved As UInteger, <MarshalAs(UnmanagedType.LPWStr)> ByVal commandstring As StringBuilder, ByVal cch As Integer) As Int32

            ' Allows client objects of the IContextMenu interface to 
            ' handle messages associated with owner-drawn menu items
            <PreserveSig()> _
            Function HandleMenuMsg(ByVal uMsg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Int32

            ' Allows client objects of the IContextMenu3 interface to 
            ' handle messages associated with owner-drawn menu items
            <PreserveSig()> _
            Function HandleMenuMsg2(ByVal uMsg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr, ByVal plResult As IntPtr) As Int32
        End Interface
    End Class
    Public Class ShellContextMenuException
        Inherits Exception
        ''' <summary>Default contructor</summary>
        Public Sub New()
        End Sub
        ''' <summary>Constructor with message</summary>
        ''' <param name="message">Message</param>
        Public Sub New(ByVal message As String)
            MyBase.New(message)
        End Sub
    End Class
    Public Class HookEventArgs
        Inherits EventArgs
        Public HookCode As Integer ' Hook code
        Public wParam As IntPtr ' WPARAM argument
        Public lParam As IntPtr ' LPARAM argument
    End Class
    Public Enum HookType As Integer
        WH_JOURNALRECORD = 0
        WH_JOURNALPLAYBACK = 1
        WH_KEYBOARD = 2
        WH_GETMESSAGE = 3
        WH_CALLWNDPROC = 4
        WH_CBT = 5
        WH_SYSMSGFILTER = 6
        WH_MOUSE = 7
        WH_HARDWARE = 8
        WH_DEBUG = 9
        WH_SHELL = 10
        WH_FOREGROUNDIDLE = 11
        WH_CALLWNDPROCRET = 12
        WH_KEYBOARD_LL = 13
        WH_MOUSE_LL = 14
    End Enum
    Public Class LocalWindowsHook
        Public Delegate Function HookProc(ByVal code As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
        Protected m_hhook As IntPtr = IntPtr.Zero
        Protected m_filterFunc As HookProc = Nothing
        Protected m_hookType As HookType
        Public Delegate Sub HookEventHandler(ByVal sender As Object, ByVal e As HookEventArgs)
        Public Event HookInvoked As HookEventHandler
        Protected Sub OnHookInvoked(ByVal e As HookEventArgs)
            RaiseEvent HookInvoked(Me, e)
        End Sub
        Public Sub New(ByVal hook As HookType)
            m_hookType = hook
            m_filterFunc = New HookProc(AddressOf Me.CoreHookProc)
        End Sub
        Public Sub New(ByVal hook As HookType, ByVal func As HookProc)
            m_hookType = hook
            m_filterFunc = func
        End Sub
        Protected Function CoreHookProc(ByVal code As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
            If code < 0 Then
                Return CallNextHookEx(m_hhook, code, wParam, lParam)
            End If

            ' Let clients determine what to do
            Dim e As New HookEventArgs()
            e.HookCode = code
            e.wParam = wParam
            e.lParam = lParam
            OnHookInvoked(e)

            ' Yield to the next hook in the chain
            Return CallNextHookEx(m_hhook, code, wParam, lParam)
        End Function
        Public Sub Install()
            m_hhook = SetWindowsHookEx(m_hookType, m_filterFunc, IntPtr.Zero, Process.GetCurrentProcess.Threads.Item(0).Id) 'CInt(AppDomain.GetCurrentThreadId())
        End Sub
        Public Sub Uninstall()
            UnhookWindowsHookEx(m_hhook)
        End Sub
        <DllImport("user32.dll")> _
        Protected Shared Function SetWindowsHookEx(ByVal code As HookType, ByVal func As HookProc, ByVal hInstance As IntPtr, ByVal threadID As Integer) As IntPtr
        End Function
        <DllImport("user32.dll")> _
        Protected Shared Function UnhookWindowsHookEx(ByVal hhook As IntPtr) As Integer
        End Function
        <DllImport("user32.dll")> _
        Protected Shared Function CallNextHookEx(ByVal hhook As IntPtr, ByVal code As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
        End Function
    End Class
    Friend Module ShellHelper
        ''' <summary>
        ''' Retrieves the High Word of a WParam of a WindowMessage
        ''' </summary>
        ''' <param name="ptr">The pointer to the WParam</param>
        ''' <returns>The unsigned integer for the High Word</returns>
        Public Function HiWord(ByVal ptr As IntPtr) As UInteger
            If (CUInt(ptr) And &H80000000UI) = &H80000000UI Then
                Return (CUInt(ptr) >> 16)
            Else
                Return CUInt((CUInt(ptr) >> 16) And &HFFFF)

            End If
        End Function

        ''' <summary>
        ''' Retrieves the Low Word of a WParam of a WindowMessage
        ''' </summary>
        ''' <param name="ptr">The pointer to the WParam</param>
        ''' <returns>The unsigned integer for the Low Word</returns>
        Public Function LoWord(ByVal ptr As IntPtr) As UInteger
            Return CUInt(CUInt(ptr) And &HFFFF)
        End Function
    End Module
End Namespace

