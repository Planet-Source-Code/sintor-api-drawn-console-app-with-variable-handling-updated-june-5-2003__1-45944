Attribute VB_Name = "modCommon"
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\/////////////////////////////////// '
' //    NICK'S VARIABLE CONSOLE APP                                   \\ '
' \\        version 1.0a                                              // '
' //        to program time: approx. 7 hours over 3 day period        \\ '
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\/////////////////////////////////// '
' //    TERMS AND CONDITIONS                                          \\ '
' \\        The code contained herein is labeled as FREEWARE and may  // '
' //        be distributed freely provided the following conditions   \\ '
' \\        are met:                                                  // '
' //            1. NO CHARGE may be issued by anyone for this code,   \\ '
' \\               with the EXCEPTION of a nominal fee for cost of    // '
' //               distribution media.                                \\ '
' \\            2. This code may NOT be implimented in ANY commercial // '
' //               setting in ANY way without express written concent \\ '
' \\               from the author.                                   // '
' //            3. All copyright and information notifications, such  \\ '
' \\               as this one, must remain intact and unmodified.    // '
' //            4. This code is provided "AS IS" with no warranty,    \\ '
' \\               expressed or otherwise, accompanying it. The       // '
' //               author is not liable for the use/misuse of this    \\ '
' \\               code.                                              // '
' //            5. Anyone who modifies this code may place their name \\ '
' \\               somewhere UNDER this notice block as a credit to   // '
' //               their work, provided you do not misrepresent       \\ '
' \\               yourself, your work, or the work of someone else.  // '
' //            6. These terms and conditions may NOT be modified in  \\ '
' \\               any way. Doing so will terminate your right to     // '
' //               use this code.                                     \\ '
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\/////////////////////////////////// '
' //    Copyright (c) 2003 Nicholas Kerr. All rights reserved.        \\ '
' \\    Use of this code requires an agreement to the terms and       // '
' //    conditions listed above.                                      \\ '
' \\        Enjoy your stay in the                                    // '
' //              realm of my code.                                   \\ '
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\/////////////////////////////////// '

Option Explicit

Private Type UcsRgbQuad         ' This code is someone else's?
 R As Byte
 G As Byte
 b As Byte
 A As Byte
End Type
Public Type POINTAPI
 x As Long
 y As Long
End Type
Public Type RECT
 Left   As Long
 Top    As Long
 Right  As Long
 Bottom As Long
End Type
Public Enum UcsPenStyles            ' So is this
 PS_SOLID = 0                ' The pen is solid.
 PS_DASH = 1                 ' The pen is dashed.
 PS_DOT = 2                  ' The pen is dotted.
 PS_DASHDOT = 3              ' The pen has alternating dashes and dots.
 PS_DASHDOTDOT = 4           ' The pen has dashes and double dots.
 PS_NULL = 5                 ' The pen is invisible.
 PS_INSIDEFRAME = 6          ' The pen is solid. When this pen is used in any GDI drawing function that takes a bounding rectangle, the dimensions of the figure are shrunk so that it fits entirely in the bounding rectangle, taking into account the width of the pen. This applies only to geometric pens.
 PS_USERSTYLE = 7            ' <b>Windows NT/2000:</b> The pen uses a styling array supplied by the user.
 PS_ALTERNATE = 8            ' <b>Windows NT/2000:</b> The pen sets every other pixel. (This style is applicable only for cosmetic pens.)
 PS_STYLE_MASK = &HF         ' Mask for previous PS_XXX values.
 PS_ENDCAP_ROUND = &H0       ' End caps are round.
 PS_ENDCAP_SQUARE = &H100    ' End caps are square.
 PS_ENDCAP_FLAT = &H200      ' End caps are flat.
 PS_ENDCAP_MASK = &HF00      ' Mask for previous PS_ENDCAP_XXX values.
 PS_JOIN_ROUND = &H0         ' Joins are beveled.
 PS_JOIN_BEVEL = &H1000      ' Joins are mitered when they are within the current limit set by the SetMiterLimit function. If it exceeds this limit, the join is beveled.
 PS_JOIN_MITER = &H2000      ' Joins are round.
 PS_JOIN_MASK = &HF000       ' Mask for previous PS_JOIN_XXX values.
 PS_COSMETIC = &H0           ' The pen is cosmetic.
 PS_GEOMETRIC = &H10000      ' The pen is geometric.
 PS_TYPE_MASK = &HF0000      ' Mask for previous PS_XXX (pen type).
End Enum
Public Enum UcsDrawTextStyles       ' And this
 DT_LEFT = &H0               ' Aligns text to the left.
 DT_TOP = &H0                ' Justifies the text to the top of the rectangle.
 DT_CENTER = &H1             ' Centers text horizontally in the rectangle.
 DT_RIGHT = &H2              ' Aligns text to the right.
 DT_VCENTER = &H64           ' Centers text vertically. This value is used only with the DT_SINGLELINE value.
 DT_BOTTOM = &H8             ' Justifies the text to the bottom of the rectangle. This value is used only with the DT_SINGLELINE value.
 DT_WORDBREAK = &H10         ' Breaks words. Lines are automatically broken between words if a word would extend past the edge of the rectangle specified by the lpRect parameter. A carriage return-line feed sequence also breaks the line.<br>If this is not specified, output is on one line.
 DT_SINGLELINE = &H20        ' Displays text on a single line only. Carriage returns and line feeds do not break the line.
 DT_EXPANDTABS = &H40        ' Expands tab characters. The default number of characters per tab is eight. The DT_WORD_ELLIPSIS, DT_PATH_ELLIPSIS, and DT_END_ELLIPSIS values cannot be used with the DT_EXPANDTABS value.
 DT_TABSTOP = &H80           ' Sets tab stops. Bits 15–8 (high-order byte of the low-order word) of the uFormat parameter specify the number of characters for each tab. The default number of characters per tab is eight. The DT_CALCRECT, DT_EXTERNALLEADING, DT_INTERNAL, DT_NOCLIP, and DT_NOPREFIX values cannot be used with the DT_TABSTOP value.
 DT_NOCLIP = &H100           ' Draws without clipping. DrawText is somewhat faster when DT_NOCLIP is used.
 DT_EXTERNALLEADING = &H200  ' Includes the font external leading in line height. Normally, external leading is not included in the height of a line of text.
 DT_CALCRECT = &H400         ' Determines the width and height of the rectangle. If there are multiple lines of text, DrawText uses the width of the rectangle pointed to by the lpRect parameter and extends the base of the rectangle to bound the last line of text. If the largest word is wider than the rectangle, the width is expanded. If the text is less than the width of the rectangle, the width is reduced. If there is only one line of text, DrawText modifies the right side of the rectangle so that it bounds the last character in the line. In either case, DrawText returns the height of the formatted text but does not draw the text.
 DT_NOPREFIX = &H800         ' Turns off processing of prefix characters. Normally, DrawText interprets the mnemonic-prefix character & as a directive to underscore the character that follows, and the mnemonic-prefix characters && as a directive to print a single &. By specifying DT_NOPREFIX, this processing is turned off
 DT_INTERNAL = &H1000        ' Uses the system font to calculate text metrics.
 DT_EDITCONTROL = &H2000     ' Duplicates the text-displaying characteristics of a multiline edit control. Specifically, the average character width is calculated in the same manner as for an edit control, and the function does not display a partially visible last line.
 DT_PATH_ELLIPSIS = &H4000   ' For displayed text, replaces characters in the middle of the string with ellipses so that the result fits in the specified rectangle. If the string contains backslash (\) characters, DT_PATH_ELLIPSIS preserves as much as possible of the text after the last backslash.<br>The string is not modified unless the DT_MODIFYSTRING flag is specified.<br>Compare with DT_END_ELLIPSIS and DT_WORD_ELLIPSIS.
 DT_END_ELLIPSIS = &H8000    ' For displayed text, if the end of a string does not fit in the rectangle, it is truncated and ellipses are added. If a word that is not at the end of the string goes beyond the limits of the rectangle, it is truncated without ellipses.<br>The string is not modified unless the DT_MODIFYSTRING flag is specified.<br>Compare with DT_PATH_ELLIPSIS and DT_WORD_ELLIPSIS.
 DT_MODIFYSTRING = &H10000   ' Modifies the specified string to match the displayed text. This value has no effect unless DT_END_ELLIPSIS or DT_PATH_ELLIPSIS is specified.
 DT_RTLREADING = &H20000     ' Layout in right-to-left reading order for bi-directional text when the font selected into the hdc is a Hebrew or Arabic font. The default reading order for all text is left-to-right.
 DT_WORD_ELLIPSIS = &H40000  ' Truncates any word that does not fit in the rectangle and adds ellipses.<br>Compare with DT_END_ELLIPSIS and DT_PATH_ELLIPSIS.
End Enum

Public Const BDR_RAISEDINNER = &H4
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENOUTER = &H2
Public Const BF_BOTTOM = &H8
Public Const BF_FLAT = &H4000      ' For flat rather than 3D borders
Public Const BF_LEFT = &H1
Public Const BF_MONO = &H8000      ' For monochrome borders.
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const EDGE_RAISED = BDR_RAISEDOUTER Or BDR_RAISEDINNER
Public Const EDGE_SUNKEN = BDR_SUNKENOUTER Or BDR_SUNKENINNER
Public Const BF_RECT = BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM

Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpPoint As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As UcsDrawTextStyles) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetForeColor Lib "gdi32" Alias "SetTextColor" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal Clr As Long, ByVal hpal As Long, ByRef lpcolorref As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Sub SetFont(ByRef sfFontSource As StdFont, ByVal sfFontDestination As StdFont)
 With sfFontDestination
  .Bold = sfFontSource.Bold
  .Charset = sfFontSource.Charset
  .Italic = sfFontSource.Italic
  .Name = sfFontSource.Name
  .Size = sfFontSource.Size
  .Strikethrough = sfFontSource.Strikethrough
  .Underline = sfFontSource.Underline
  .Weight = sfFontSource.Weight
 End With
End Sub

Public Function comCreatePen(ByVal transColor As OLE_COLOR, Optional ByVal penWidth As Long = 1, Optional ByRef penStyle As UcsPenStyles = UcsPenStyles.PS_SOLID) As Long
Dim cCol As Long
 Call OleTranslateColor(transColor, 0&, cCol)
 comCreatePen = CreatePen(penStyle, penWidth, cCol)
End Function

' This function is someone else's
Public Function ColorGrad(ByVal Color1 As Long, ByVal Color2 As Long, ByVal Steps As Long) As Long()
Dim b As UcsRgbQuad, n As UcsRgbQuad, o As UcsRgbQuad
Dim factR As Single, factG As Single, factB As Single
Dim i As Long, j As Long
Dim cGrad() As Long
 j = (Steps - 1)
 ReDim cGrad(0 To j)
 ClrBit Color1, n
 ClrBit Color2, o
 factR = ((CLng(o.R) - CLng(n.R)) / (Steps - 1))
 factG = ((CLng(o.G) - CLng(n.G)) / (Steps - 1))
 factB = ((CLng(o.b) - CLng(n.b)) / (Steps - 1))
 For i = 0 To j
  b.R = n.R + (factR * i)
  b.G = n.G + (factG * i)
  b.b = n.b + (factB * i)
  With b
   cGrad(i) = (.R) + (.G * &H100&) + (.b * &H10000)
  End With
 Next i
 ColorGrad = cGrad
End Function

' So is this
Private Sub ClrBit(ByVal Color As Long, Bits As UcsRgbQuad)
 GetRGB Color, Bits.R, Bits.G, Bits.b
End Sub

' And this
Private Sub GetRGB(ByVal Color As Long, Red As Byte, Green As Byte, Blue As Byte)
Dim c As Long
 c = (Color And &HFF&)
 Red = CByte(c)
 c = ((Color And &HFF00&) / &H100&)
 Green = CByte(c)
 c = ((Color And &HFF0000) / &H10000)
 Blue = CByte(c)
End Sub

Public Function CreateBrush(ByVal nColor As OLE_COLOR) As Long
Dim cCol As Long
 Call OleTranslateColor(nColor, 0&, cCol)
 CreateBrush = CreateSolidBrush(cCol)
End Function

' This one definately isn't mine
Public Function FormatCount(Count As Long, Optional FormatType As Byte = 0) As String
Dim Days As Integer, Hours As Long, Minutes As Long, Seconds As Long, Miliseconds As Long
 Miliseconds = Count Mod 1000
 Count = Count \ 1000
 Days = Count \ (24& * 3600&)
 If Days > 0 Then Count = Count - (24& * 3600& * Days)
 Hours = Count \ 3600&
 If Hours > 0 Then Count = Count - (3600& * Hours)
 Minutes = Count \ 60
 Seconds = Count Mod 60
 Select Case FormatType
  Case 0: FormatCount = Days & " dd, " & Hours & " h, " & Minutes & " min, " & Seconds & " s, " & Miliseconds & " ms"
  Case 1: FormatCount = Days & " days, " & Hours & " hours, " & Minutes & " minutes, " & Seconds & " seconds, " & Miliseconds & " miliseconds"
  Case 2: FormatCount = Days & ":" & Hours & ":" & Minutes & ":" & Seconds & ":" & Miliseconds
 End Select
End Function

Public Function SetTextColor(ByVal hDC As Long, ByVal crColor As Long) As Long
Dim tCol As Long
 Call OleTranslateColor(crColor, 0&, tCol)
 SetTextColor = SetForeColor(hDC, tCol)
End Function

' Footnote: I appologize to whomever owns some of the code I've implimented
'           in this module. I've had it so long I've misplaced the author's names.
