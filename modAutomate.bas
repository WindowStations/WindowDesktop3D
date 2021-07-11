Attribute VB_Name = "modAutomate"
'ThisWorkbook.VBProject.References.AddFromFile "C:\Windows\SysWOW64\UIAutomationCore.dll"
'Public Type POINTAPI
'    x As Long
'    y As Long
'End Type
'Public Declare Function apiGetCursorPos Lib "user32" Alias "GetCursorPos" (ByRef lpPoint As POINTAPI) As Long
'Public Declare Function apiElementFromPoint Lib "UIAutomationCore" Alias "ElementFromPoint" (ByRef tp As tagPOINT) As IUIAutomationElement
Public Function InvokeElement(ByVal hWnd As Long, ByVal sText As String) As Boolean
   On Error GoTo poop
   InvokeElement = False
   If hWnd = 0 Then Exit Function
   Dim uAuto   As IUIAutomation
   Dim el      As IUIAutomationElement
   Dim uCond   As IUIAutomationCondition
   Dim elName  As IUIAutomationElement
   Dim pInvoke As IUIAutomationInvokePattern
   Set uAuto = New CUIAutomation
   Set uCond = uAuto.CreatePropertyCondition(UIA_NamePropertyId, sText)
   If uCond Is Nothing Then Exit Function
   Set el = uAuto.ElementFromHandle(ByVal hWnd)
   If el Is Nothing Then Exit Function
   Set elName = el.FindFirst(TreeScope_Children, uCond)
   If elName Is Nothing Then Exit Function
   Set pInvoke = elName.GetCurrentPattern(UIA_InvokePatternId)
   If pInvoke Is Nothing Then Exit Function
   pInvoke.Invoke
   InvokeElement = True
poop:
End Function
Public Function InvokeElement3(ByVal el As IUIAutomationElement, ByVal sText As String) As Boolean
   On Error GoTo poop
   InvokeElement3 = False
   Dim uAuto   As IUIAutomation
   ' Dim el As IUIAutomationElement
   Dim uCond   As IUIAutomationCondition
   Dim elName  As IUIAutomationElement
   Dim pInvoke As IUIAutomationInvokePattern
   Set uAuto = New CUIAutomation
   Set uCond = uAuto.CreatePropertyCondition(UIA_NamePropertyId, sText)
   If uCond Is Nothing Then Exit Function
   'Set el = uAuto.ElementFromHandle(ByVal hwnd)
   ' If el Is Nothing Then Exit Function
   Set elName = el.FindFirst(TreeScope_Children, uCond)
   If elName Is Nothing Then Exit Function
   Set pInvoke = elName.GetCurrentPattern(UIA_InvokePatternId)
   If pInvoke Is Nothing Then Exit Function
   pInvoke.Invoke
   InvokeElement3 = True
poop:
End Function
Public Function GetUIADesktopElements() As IUIAutomationElementArray
   On Error GoTo poop
   Dim allChilds As IUIAutomationElementArray
   Set GetUIADesktopElements = allChilds
   Dim oUIAutomation As New CUIAutomation
   Dim oUIADesktop   As IUIAutomationElement
   Set oUIADesktop = oUIAutomation.GetRootElement
   If oUIADesktop Is Nothing Then Exit Function
   Set allChilds = oUIADesktop.FindAll(TreeScope_Children, oUIAutomation.CreateTrueCondition)
   Set GetUIADesktopElements = allChilds
poop:
End Function
Public Function GetUIAChildElements(ByVal el As IUIAutomationElement) As IUIAutomationElementArray
   On Error GoTo poop
   Dim allChilds As IUIAutomationElementArray
   Set GetUIAChildElements = allChilds
   Dim oUIAutomation As New CUIAutomation
   Set allChilds = el.FindAll(TreeScope_Children, oUIAutomation.CreateTrueCondition)
   Set GetUIAChildElements = allChilds
poop:
End Function
Public Function GetUIADesktopIconElements() As IUIAutomationElementArray
   On Error GoTo poop
   Dim oUIAutomation As New CUIAutomation
   Dim oUIADesktop   As IUIAutomationElement
   Dim elpm          As IUIAutomationElement
   Dim allChilds     As IUIAutomationElementArray
   Dim uCond1        As IUIAutomationCondition
   Dim uCond2        As IUIAutomationCondition
   Dim ucond3        As IUIAutomationCondition
   Set GetUIADesktopIconElements = allChilds
   Set oUIADesktop = oUIAutomation.GetRootElement
   If oUIADesktop Is Nothing Then GoTo poop
   Set uCond1 = oUIAutomation.CreatePropertyCondition(UIA_ClassNamePropertyId, "Progman")
   Set uCond2 = oUIAutomation.CreatePropertyCondition(UIA_NamePropertyId, "Program Manager")
   Set ucond3 = oUIAutomation.CreateOrCondition(uCond1, uCond2)
   Set elpm = oUIADesktop.FindFirst(TreeScope_Children, ucond3)
   If elpm Is Nothing Then GoTo poop
   Set uCond1 = oUIAutomation.CreatePropertyCondition(UIA_ClassNamePropertyId, "SysListView32")
   Set allChilds = elpm.FindAll(TreeScope_Descendants, uCond1)
   If allChilds.Length = 1 Then
      Set allChilds = allChilds.GetElement(0).FindAll(TreeScope_Descendants, oUIAutomation.CreateTrueCondition)
      Set GetUIADesktopIconElements = allChilds
   End If
   Exit Function
poop:
   MsgBox Err.Description
End Function
Public Sub InvoketUIADesktopIconElement(ByVal index As Long)
   On Error GoTo poop
   Dim oUIAutomation As New CUIAutomation
   Dim oUIADesktop   As IUIAutomationElement
   Dim elpm          As IUIAutomationElement
   Dim allChilds     As IUIAutomationElementArray
   Dim el            As IUIAutomationElement
   Dim uCond1        As IUIAutomationCondition
   Dim uCond2        As IUIAutomationCondition
   Dim ucond3        As IUIAutomationCondition
   Set oUIADesktop = oUIAutomation.GetRootElement
   If oUIADesktop Is Nothing Then GoTo poop
   Set uCond1 = oUIAutomation.CreatePropertyCondition(UIA_ClassNamePropertyId, "Progman")
   Set uCond2 = oUIAutomation.CreatePropertyCondition(UIA_NamePropertyId, "Program Manager")
   Set ucond3 = oUIAutomation.CreateOrCondition(uCond1, uCond2)
   Set elpm = oUIADesktop.FindFirst(TreeScope_Children, ucond3) '
   If elpm Is Nothing Then GoTo poop
   Set uCond1 = oUIAutomation.CreatePropertyCondition(UIA_ClassNamePropertyId, "SysListView32")
   Set allChilds = elpm.FindAll(TreeScope_Descendants, uCond1)
   If allChilds.Length = 1 Then
      Set allChilds = allChilds.GetElement(0).FindAll(TreeScope_Descendants, oUIAutomation.CreateTrueCondition)
      InvokeElement2 allChilds.GetElement(index)
   End If
   Exit Sub
poop:
   MsgBox Err.Description
End Sub
Public Sub InvoketUIADesktopIconElementByName(ByVal name As String)
   On Error GoTo poop
   Dim oUIAutomation As New CUIAutomation
   Dim oUIADesktop   As IUIAutomationElement
   Dim elpm          As IUIAutomationElement
   Dim allChilds     As IUIAutomationElementArray
   Dim el            As IUIAutomationElement
   Dim i             As Long
   Dim uCond1        As IUIAutomationCondition
   Dim uCond2        As IUIAutomationCondition
   Dim ucond3        As IUIAutomationCondition
   Set oUIADesktop = oUIAutomation.GetRootElement
   If oUIADesktop Is Nothing Then GoTo poop
   Set uCond1 = oUIAutomation.CreatePropertyCondition(UIA_ClassNamePropertyId, "Progman")
   Set uCond2 = oUIAutomation.CreatePropertyCondition(UIA_NamePropertyId, "Program Manager")
   Set ucond3 = oUIAutomation.CreateOrCondition(uCond1, uCond2)
   Set elpm = oUIADesktop.FindFirst(TreeScope_Children, ucond3) '
   If elpm Is Nothing Then GoTo poop
   Set uCond1 = oUIAutomation.CreatePropertyCondition(UIA_ClassNamePropertyId, "SysListView32")
   Set allChilds = elpm.FindAll(TreeScope_Descendants, uCond1)
   If allChilds.Length = 1 Then
      Set allChilds = allChilds.GetElement(0).FindAll(TreeScope_Descendants, oUIAutomation.CreateTrueCondition)
      For i = 0 To allChilds.Length - 1
         If allChilds.GetElement(i).CurrentName = name Then
            InvokeElement2 allChilds.GetElement(i)
            Beep
            Exit For
         End If
      Next
   End If
   Exit Sub
poop:
   MsgBox Err.Description
End Sub
Public Function InvokeElement2(ByVal el As IUIAutomationElement) As Boolean
   On Error GoTo poop
   Dim pInvoke As IUIAutomationInvokePattern
   Set pInvoke = el.GetCurrentPattern(UIA_InvokePatternId)
   If pInvoke Is Nothing Then Exit Function
   pInvoke.Invoke
   InvokeElement2 = True
poop:
End Function
Public Function GetUIAAncestorElements(ByVal el As IUIAutomationElement) As IUIAutomationElementArray
   On Error GoTo poop
   Dim allChilds As IUIAutomationElementArray
   Set GetUIAAncestorElements = allChilds
   Dim oUIAutomation As New CUIAutomation
   Set allChilds = el.FindAll(TreeScope_Ancestors, oUIAutomation.CreateTrueCondition)
   Set GetUIAAncestorElements = allChilds
poop:
End Function
Public Function GetUIAAncestorElement(ByRef el As IUIAutomationElement, ByVal classname As String, ByVal name As String) As IUIAutomationElement
   On Error GoTo poop
   Dim allChilds     As IUIAutomationElementArray
   Dim oUIAutomation As New CUIAutomation
   'Dim uCond1 As IUIAutomationCondition
   'Dim uCond2 As IUIAutomationCondition
   'Set uCond1 = uAuto.CreatePropertyCondition(UIA_ClassNamePropertyId, classname)
   ' Set uCond2 = uAuto.CreatePropertyCondition(UIA_NamePropertyId, name)
   Set allChilds = el.FindAll(TreeScope_Ancestors, oUIAutomation.CreateTrueCondition)
   MsgBox allChilds.Length
   Set GetUIAAncestorElement = allChilds
   Exit Function
poop:
   MsgBox Err.Description
End Function
Public Function WalkUIADesktopElements() As IUIAutomationElement()
   On Error GoTo poop
   Dim el()          As IUIAutomationElement
   Dim oUIAutomation As New CUIAutomation
   Dim oUIADesktop   As IUIAutomationElement
   Dim allChilds     As IUIAutomationElementArray
   Dim oUIElement    As IUIAutomationElement
   Dim oTW           As IUIAutomationTreeWalker
   Set oUIADesktop = oUIAutomation.GetRootElement
   If oUIADesktop Is Nothing Then Exit Function
   Set oTW = oUIAutomation.ControlViewWalker
   Set oUIElement = oTW.GetFirstChildElement(oUIADesktop)
   Dim i As Long
   Do
      If oUIElement Is Nothing Then Exit Do ' exit loop
      On Error GoTo skip
      ReDim Preserve el(i)
      Set el(i) = oUIElement
skip:
      Set oUIElement = oTW.GetNextSiblingElement(oUIElement)
      i = i + 1
   Loop
poop:
   WalkUIADesktopElements = el
End Function
Public Function ProcessIdToHandleUIA(ByVal pid As Long) As Long
   ProcessIdToHandleUIA = 0
   On Error GoTo poop
   Dim el()          As IUIAutomationElement
   Dim oUIAutomation As New CUIAutomation
   Dim oUIADesktop   As IUIAutomationElement
   Dim allChilds     As IUIAutomationElementArray
   Dim oUIElement    As IUIAutomationElement
   Dim oTW           As IUIAutomationTreeWalker
   Set oUIADesktop = oUIAutomation.GetRootElement
   If oUIADesktop Is Nothing Then Exit Function
   Set oTW = oUIAutomation.ControlViewWalker
   Set oUIElement = oTW.GetFirstChildElement(oUIADesktop)
   ' auwlk = Windows.Automation.TreeWalker.ControlViewWalker.GetFirstChild(AutomationElement.RootElement)
   Dim i As Long
   For i = 1 To 1000
      On Error GoTo skip
      If oUIElement Is Nothing Then Exit For
      Dim hWnd As Long
      hWnd = GetUIAHandle(oUIElement)
      '        If hwnd <> 0 And pid = GetWindowProcessId(hwnd) Then
      '            ProcessIdToHandleUIA = hwnd: Exit Function
      '        End If
      '        oUIElement = oTW.GetNextSibling(oUIElement)
skip:
   Next
poop:
End Function
Public Function GetUIAHandle(ByVal el As IUIAutomationElement) As Long
   Dim hWnd As Long: hWnd = 0
   GetUIAHandle = hWnd
   If el Is Nothing Then Exit Function
   On Error GoTo skip
   hWnd = el.GetCurrentPropertyValue(UIA_NativeWindowHandlePropertyId)
skip:
   GetUIAHandle = hWnd
End Function
Public Function GetUIAProcessId(ByVal el As IUIAutomationElement) As Long
   Dim pid As Long: pid = 0
   GetUIAProcessId = pid
   If el Is Nothing Then Exit Function
   On Error GoTo skip
   pid = el.GetCurrentPropertyValue(UIA_ProcessIdPropertyId)
skip:
   GetUIAProcessId = pid
End Function
'Public Function GetUIAControlType(ByVal el As IUIAutomationElement) As Long
'    Dim hwnd As Long: hwnd = 0
'    GetUIAControlType = hwnd
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    hwnd = el.GetCurrentPropertyValue(UIA_ControlTypePropertyId)
'skip:
'    GetUIAControlType = hwnd
'End Function
Public Function GetUIALocalizedControlType(ByVal el As IUIAutomationElement) As String
   Dim lct As String: lct = ""
   GetUIALocalizedControlType = lct
   If el Is Nothing Then Exit Function
   On Error GoTo skip
   lct = el.GetCurrentPropertyValue(UIA_LocalizedControlTypePropertyId)
skip:
   GetUIALocalizedControlType = lct
End Function
Public Function GetUIANameProperty(ByVal el As IUIAutomationElement) As String
   Dim name As String: name = ""
   GetUIANameProperty = name
   If el Is Nothing Then Exit Function
   On Error GoTo skip
   name = el.GetCurrentPropertyValue(UIA_NamePropertyId)
skip:
   GetUIANameProperty = name
End Function
Public Function GetUIAClassName(ByVal el As IUIAutomationElement) As String
   Dim cname As String: cname = ""
   GetUIAClassName = cname
   If el Is Nothing Then Exit Function
   On Error GoTo skip
   cname = el.GetCurrentPropertyValue(UIA_ClassNamePropertyId)
skip:
   GetUIAClassName = cname
End Function
Public Function GetUIAFullDescription(ByVal el As IUIAutomationElement) As String
   Dim desc As String: desc = ""
   GetUIAFullDescription = desc
   If el Is Nothing Then Exit Function
   On Error GoTo skip
   desc = el.GetCurrentPropertyValue(UIA_FullDescriptionPropertyId)
skip:
   GetUIAFullDescription = desc
End Function
Public Function GetUIAHelpText(ByVal el As IUIAutomationElement) As String
   Dim help As String: help = ""
   GetUIAHelpText = help
   If el Is Nothing Then Exit Function
   On Error GoTo skip
   help = el.GetCurrentPropertyValue(UIA_HelpTextPropertyId)
skip:
   GetUIAHelpText = help
End Function
Public Function GetUIAHasKeyboardFocus(ByVal el As IUIAutomationElement) As Boolean
   Dim b As Boolean: b = False
   GetUIAHasKeyboardFocus = b
   If el Is Nothing Then Exit Function
   On Error GoTo skip
   b = el.GetCurrentPropertyValue(UIA_HasKeyboardFocusPropertyId)
skip:
   GetUIAHasKeyboardFocus = b
End Function
Public Function GetUIAIsEnabled(ByVal el As IUIAutomationElement) As Boolean
   Dim b As Boolean: b = False
   GetUIAIsEnabled = b
   If el Is Nothing Then Exit Function
   On Error GoTo skip
   b = el.GetCurrentPropertyValue(UIA_IsEnabledPropertyId)
skip:
   GetUIAIsEnabled = b
End Function
Public Function GetUIAIsKeyboardFocusable(ByVal el As IUIAutomationElement) As Boolean
   Dim b As Boolean: b = False
   GetUIAIsKeyboardFocusable = b
   If el Is Nothing Then Exit Function
   On Error GoTo skip
   b = el.GetCurrentPropertyValue(UIA_IsKeyboardFocusablePropertyId)
skip:
   GetUIAIsKeyboardFocusable = b
End Function
Public Function GetUIAIsControlElement(ByVal el As IUIAutomationElement) As Boolean
   Dim b As Boolean: b = False
   GetUIAIsControlElement = b
   If el Is Nothing Then Exit Function
   On Error GoTo skip
   b = el.GetCurrentPropertyValue(UIA_IsControlElementPropertyId)
skip:
   GetUIAIsControlElement = b
End Function
Public Function GetUIAIsContentElement(ByVal el As IUIAutomationElement) As Boolean
   Dim b As Boolean: b = False
   GetUIAIsContentElement = b
   If el Is Nothing Then Exit Function
   On Error GoTo skip
   b = el.GetCurrentPropertyValue(UIA_IsContentElementPropertyId)
skip:
   GetUIAIsContentElement = b
End Function
Public Function GetUIAIsOffscreen(ByVal el As IUIAutomationElement) As Boolean
   Dim b As Boolean: b = False
   GetUIAIsOffscreen = b
   If el Is Nothing Then Exit Function
   On Error GoTo skip
   b = el.GetCurrentPropertyValue(UIA_IsOffscreenPropertyId)
skip:
   GetUIAIsOffscreen = hWnd
End Function
Public Function GetUIAIsPassword(ByVal el As IUIAutomationElement) As Boolean
   Dim b As Boolean: b = False
   GetUIAIsPassword = b
   If el Is Nothing Then Exit Function
   On Error GoTo skip
   b = el.GetCurrentPropertyValue(UIA_IsPasswordPropertyId)
skip:
   GetUIAIsPassword = b
End Function
'Public Function GetUIABoundingRectangle(ByVal el As IUIAutomationElement) As Variant
'    Dim hwnd As Variant
'    GetUIABoundingRectangle = hwnd
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    hwnd = el.GetCurrentPropertyValue(UIA_BoundingRectanglePropertyId)
'skip:
'    GetUIABoundingRectangle = hwnd
'End Function
'Public Function GetUIASize(ByVal el As IUIAutomationElement) As Long
'    Dim hwnd As Long: hwnd = 0
'    GetUIASize = hwnd
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    hwnd = el.GetCurrentPropertyValue(UIA_SizePropertyId)
'skip:
'    GetUIASize = hwnd
'End Function
'Public Function GetUIAClickablePoint(ByVal el As IUIAutomationElement) As Long
'    Dim hwnd As Long: hwnd = 0
'    GetUIAClickablePoint = hwnd
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    hwnd = el.GetCurrentPropertyValue(UIA_ClickablePointPropertyId)
'skip:
'    GetUIAClickablePoint = hwnd
'End Function
'Public Function GetUIACenterPoint(ByVal el As IUIAutomationElement) As Long
'    Dim hwnd As Long: hwnd = 0
'    GetUIACenterPoint = hwnd
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    hwnd = el.GetCurrentPropertyValue(UIA_CenterPointPropertyId)
'skip:
'    GetUIACenterPoint = hwnd
'End Function
Public Function GetUIAItemStatus(ByVal el As IUIAutomationElement) As Long
   Dim itemstat As Long
   GetUIAItemStatus = itemstat
   If el Is Nothing Then Exit Function
   On Error GoTo skip
   itemstat = el.GetCurrentPropertyValue(UIA_ItemStatusPropertyId)
skip:
   GetUIAItemStatus = itemstat
End Function
Public Function GetUIAItemType(ByVal el As IUIAutomationElement) As Long
   Dim hWnd As Long: hWnd = 0
   GetUIAItemType = hWnd
   If el Is Nothing Then Exit Function
   On Error GoTo skip
   hWnd = el.GetCurrentPropertyValue(UIA_ItemTypePropertyId)
skip:
   GetUIAItemType = hWnd
End Function
'
'
'
'
'
'
Public Function GetElementFromPoint() As IUIAutomationElement
   '    On Error GoTo poop
   '    Dim UI As New UIAutomationClient.CUIAutomation
   '    Dim El As UIAutomationClient.IUIAutomationElement
   '    Dim tag As UIAutomationClient.tagPOINT
   '    Set GetElementFromPoint = El
   '    Set UI = New CUIAutomation
   '    Dim p As POINTAPI
   '    If apiGetCursorPos(p) = 0 Then Exit Function
   '    tag.x = p.x
   '    tag.y = p.y
   '    Set El = UI.ElementFromPoint(tag)
   '    Set GetElementFromPoint = El
'poop:
End Function
Public Function MainWindowHandles(ByVal pid As Long) As Long()
   Dim hwnds() As Long
   '    Dim c() As Long
   '    c = ChildWindows(0)
   '    Dim n As Long: n = 0
   '    Dim i As Long: i = 0
   '    Do
   '        On Error GoTo skip
   '        Dim id As Long
   '        id = GetPIDFromHWND(c(i))
   '        If id = pid Then
   '            MsgBox id & " " & c(i)
   '
   '            ReDim Preserve hwnds(n)
   '            hwnds(n) = c(i)
   '            n = n + 1
   '        End If
'skip:
   '        If i = UBound(c) Then Exit Do
   '        i = i + 1
   '    Loop
   MainWindowHandles = hwnds
End Function
Public Function MainWindowTitles(ByVal pid As Long) As String()
   Dim titles() As String
   '    Dim c() As Long
   '    c = ChildWindows(0)
   '    Dim n As Long: n = 0
   '    Dim i As Long: i = 0
   '    For i = 0 To UBound(c)
   '        On Error GoTo skip
   '        Dim id As Long
   '        id = GetPIDFromHWND(c(i))
   '        If id = pid Then
   '            ReDim Preserve titles(n)
   '            ' titles(n) = c(i)
   '            n = n + 1
   '        End If
'skip:
   '    Next
   MainWindowTitles = titles
End Function
Public Sub CloseMainWindows(ByVal pid As Long)
   '    Dim c() As Long
   '    c = ChildWindows(0)
   '    Dim n As Long: n = 0
   '    Dim i As Long: i = 0
   '    For i = 0 To UBound(c)
   '        On Error GoTo skip
   '        Dim id As Long
   '        id = GetPIDFromHWND(c(i))
   '        If id = pid Then
   '            'apiPostMessage(c(i), WM_CLOSE, 0,0)
   '            n = n + 1
   '        End If
'skip:
   '    Next
End Sub
