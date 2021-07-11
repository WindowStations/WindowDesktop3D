Attribute VB_Name = "modTextFit"
Option Explicit
Public Enum qeFitPictureAlign
   eLeft
   eCentre
   eRight
   eJustify
End Enum
Public Enum qeFitPictureChar
   eNone
   eSpace
   eDash
   eLine
   eOops
End Enum
Public Enum qeFitPictureShadow
   eTopLeft
   eTop
   eTopRight
   eLeft
   eNoShadow
   eRight
   eBottomLeft
   eBottom
   eBottomRight
End Enum
Private Type qtFitPictureLine
   sLine As String
   eEnd As qeFitPictureChar
End Type
Public Function TextToPicture(Picture As PictureBox, sString As String, eAlign As qeFitPictureAlign, Optional sBorder As Single = 60, Optional eShadow As qeFitPictureShadow = eNoShadow, Optional lShadowColor As Long = vb3DShadow) As Boolean
   Dim tLine() As qtFitPictureLine
   Dim iLine As Integer
   Dim iCount As Integer
   Dim iFont As Integer
   Dim iSpace As Integer
   Dim iMarker As Integer
   Dim sSizeX As Single
   Dim sSizeY As Single
   Dim sHeight As Single
   Dim sWidth As Single
   Dim sArea As Single
   Dim sLineHeight As Single
   Dim sCharWidth As Single
   Dim sWord As String
   Dim sChar As String
   Dim eCharType As qeFitPictureChar
   Dim bNewLine As Boolean
   Dim bFound As Boolean
   Dim sOffsetX As Single
   Dim sOffsetY As Single
   Dim lForeColor As Long
   On Error GoTo TextToPictureError
   iSpace = StringCount(sString, vbCrLf) ' Find Carriage Line break (vbCrLf) characters
   With Picture
      If sBorder * 2 > .ScaleWidth Then
         GoTo TextToPictureError ' BORDER CHECK: Wider than the width of the picture
      End If
      If sBorder * 2 > .ScaleHeight Then ' BORDER CHECK: Taller than the height of the picture
         Stop
      End If
      sWidth = .ScaleWidth - sBorder * 2 ' BORDER CALCULATE: Dimensions of box minus border
      sHeight = .ScaleHeight - sBorder * 2
      sArea = sWidth * sHeight ' FONT SIZE: Estimate likely fontsize (slight over-estimation)
      iCount = 6
      Do
         .FontSize = iCount
         sSizeX = .TextWidth(sString)
         sSizeY = .TextHeight(" ") ' NEXT LINE: Estimate the line height (including the number - of line breaks calculated above)
         sLineHeight = ((sSizeX / sWidth) + iSpace) * sSizeY ' SIZE CHECK: Check size or increase font size
         If sLineHeight >= sHeight Then
            bFound = True
         Else
            iFont = iCount
         End If
         iCount = iCount + 1
      Loop While Not bFound And iFont < 72
      If iFont = 0 Then ' FONT CHECK: Was a valid fontsize found
         GoTo TextToPictureError ' FONT CHECK: Text to large
      End If
Do ' LINE SPLIT: Cut text to line width
         .FontSize = iFont
         iCount = 1
         iLine = 1
         ReDim tLine(1)
         sWord = ""
         Do
            Do
               eCharType = eNone
               sChar = Mid$(sString, iCount, 1) 'CHARACTER CHECK: Look for potential line breaks or where text width is greater than boundary
               

               
               
               Select Case sChar
               Case " "
                  eCharType = eSpace
               Case "-"
                  sSizeX = .TextWidth(tLine(iLine).sLine & sWord & sChar)
                  If sSizeX > sWidth Then
                     eCharType = eOops
                  Else
                     eCharType = eDash
                  End If
               Case vbLf
                  sChar = ""
                  eCharType = eLine
               Case vbCr
                  If iCount < Len(sString) Then
                     If Mid$(sString, iCount + 1, 1) = vbLf Then
                        iCount = iCount + 1
                     End If
                  End If
                  sChar = ""
                  eCharType = eLine
               Case Else
                  sSizeX = .TextWidth(tLine(iLine).sLine & sWord & sChar) ' CHARACTER CHECK: See if addition of character makes line too long
                  If sSizeX > sWidth Then
                     eCharType = eOops
                  Else
                     sWord = sWord & sChar
                  End If
               End Select
               iCount = iCount + 1
            Loop While iCount <= Len(sString) And eCharType = eNone
            bNewLine = False ' LINE SPLIT: Examine potential line break
            Select Case eCharType
            Case qeFitPictureChar.eNone
               tLine(iLine).sLine = tLine(iLine).sLine & sWord
               tLine(iLine).eEnd = eLine
            Case qeFitPictureChar.eOops
               If tLine(iLine).eEnd = eNone Then
                  tLine(iLine).sLine = sWord
                  sWord = sChar
               Else
                  tLine(iLine).sLine = Trim$(tLine(iLine).sLine)
                  sWord = sWord & sChar
               End If
               bNewLine = True
            Case qeFitPictureChar.eDash, qeFitPictureChar.eSpace
               tLine(iLine).eEnd = eCharType
               tLine(iLine).sLine = tLine(iLine).sLine & sWord & sChar
               sWord = ""
            Case qeFitPictureChar.eLine
               tLine(iLine).sLine = tLine(iLine).sLine & sWord
               tLine(iLine).eEnd = eLine
               sWord = ""
               bNewLine = True
            End Select
            If bNewLine Then ' LINE SPLIT: Add new line if required
               iLine = iLine + 1
               ReDim Preserve tLine(iLine)
            End If
         Loop While iCount <= Len(sString)
         bFound = CBool(iLine * .TextHeight("X") > sHeight) ' TEXT FIT: Check the height is acceptable
         If bFound Then
            iFont = iFont - 1  ' TEXT FIT: Font size is too large - decrease and try again
         End If
      Loop While bFound
      sOffsetX = ((eShadow Mod 3) - 1) * (Screen.TwipsPerPixelX * ((iFont \ 15) + 1)) ' SHADOW: Calculate position of shadow offset
      sOffsetY = ((eShadow \ 3) - 1) * (Screen.TwipsPerPixelY * ((iFont \ 15) + 1))
      lForeColor = .ForeColor
      If eShadow <> eNoShadow Then
         .ForeColor = lShadowColor
      End If
      Do
         iCount = 1
         .CurrentY = sBorder + sOffsetY
         Do
            .CurrentX = sBorder + sOffsetX
            tLine(iCount).sLine = VBA.Trim(tLine(iCount).sLine)
            Select Case eAlign ' ALIGNMENT: Calculate position of line dependent on alignment setting
            Case qeFitPictureAlign.eLeft
               Picture.Print tLine(iCount).sLine
            Case qeFitPictureAlign.eCentre
               sSizeX = (sWidth - .TextWidth(tLine(iCount).sLine)) / 2 + sBorder
               .CurrentX = sSizeX + sOffsetX
               Picture.Print tLine(iCount).sLine
            Case qeFitPictureAlign.eRight
               sSizeX = sWidth - .TextWidth(tLine(iCount).sLine) + sBorder
               .CurrentX = sSizeX + sOffsetX
               Picture.Print tLine(iCount).sLine
            Case qeFitPictureAlign.eJustify
               If tLine(iCount).eEnd <> eLine Then ' ' ALIGNMENT: Full justification is more complex.  Find spacesand calculate extra spacing required
                  sCharWidth = .TextWidth(" ") ' NEXT LINE: Check to see if line has an line break
                  iSpace = 0
                  iMarker = 0
                  Do
                     iMarker = InStr(iMarker + 1, tLine(iCount).sLine, " ")
                     If iMarker > 0 Then
                        iSpace = iSpace + 1
                     End If
                  Loop While iMarker > 0
                  sSizeX = sWidth - .TextWidth(tLine(iCount).sLine)
                  bFound = False
                  If iSpace > 0 Then ' ALIGNMENT: Check number of spaces and extra size, if too large  use character justification as well as word justification
                     If sSizeX \ iSpace > sCharWidth * 3 Then
                        bFound = True
                     End If
                  Else
                     bFound = True
                  End If
                  If bFound Then
                     sSizeY = Len(tLine(iCount).sLine) - 1 + (iSpace * 2)
                     sSizeY = sSizeX / sSizeY
                     sSizeX = sSizeY * 3
                  Else
                     sSizeX = sSizeX / iSpace
                     sSizeY = 0
                  End If
                  iMarker = 1
                  Do While iMarker <= Len(tLine(iCount).sLine)
                     sChar = Mid$(tLine(iCount).sLine, iMarker, 1)
                     sCharWidth = .CurrentX + .TextWidth(sChar)
                     sLineHeight = .CurrentY
                     Picture.Print sChar
                     If sChar = " " Then
                        sCharWidth = sCharWidth + sSizeX
                     Else
                        sCharWidth = sCharWidth + sSizeY
                     End If
                     .CurrentX = sCharWidth
                     .CurrentY = sLineHeight
                     iMarker = iMarker + 1
                  Loop
                  Picture.Print ""
               Else
                  Picture.Print tLine(iCount).sLine
               End If
            End Select
            iCount = iCount + 1
         Loop While iCount <= iLine
         If .ForeColor <> lForeColor Then ' SHADOW: Check current status of shadow repeat print process ifrequired
            .ForeColor = lForeColor
            sOffsetX = 0
            sOffsetY = 0
         Else
            eShadow = eNoShadow
         End If
      Loop While eShadow <> eNoShadow
   End With
   TextToPicture = True
   Exit Function
TextToPictureError:
   TextToPicture = False ' ERROR: Could not display text in picture
End Function
Private Function StringCount(ByVal Expression As String, Item As String) As Integer
   On Error Resume Next
   Dim lPosition As Integer
   Dim lCount As Integer
   Do
      lPosition = InStr(lPosition + 1, Expression, Item)
      If lPosition > 0 Then
         lCount = lCount + 1
      End If
   Loop While lPosition > 0
   StringCount = lCount
End Function
