VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const Prompt = "]"
Private Const RET_LINE_CONST = 35   ' #

Private client  As Object
Private tWidth  As Long
Private tHeight As Long
Private cRevLn  As Long
Private cConsoleText()  As Variant
Private CurrentLineText As Variant

Public Sub Remap(Optional ByVal cText As String = "X")
 tWidth = client.TextWidth(cText)
 tHeight = client.TextHeight(cText)
End Sub

Public Property Set Target(ByRef cVal As Object)
 If TypeOf cVal Is Form Then
  Set client = cVal
  Call cVarC.CreateVar("cvar_backCol", 0)
  Call cVarC.CreateVar("cvar_foreNorm", 16777215)
  Call cVarC.CreateVar("cvar_foreResp", 16761024)
  Call cVarC.CreateVar("Nick_Kerr", "The author of this program!", , True)
  Call CreateTimer(hwnd, 1000)
  Call Remap
  Call DrawConsole
 End If
End Property

Public Property Get gRLC() As Long
 gRLC = RET_LINE_CONST
End Property

Public Property Get hwnd() As Long
 hwnd = client.hwnd
End Property

Public Property Get hDC() As Long
 hDC = client.hDC
End Property

Public Sub AddLine(ByVal cLine As Variant)
 cConsoleText(UBound(cConsoleText)) = cLine
 ReDim Preserve cConsoleText(UBound(cConsoleText) + 1)
End Sub

Public Sub AddCharToLine(ByVal ChrID As Integer)
Dim tT As String
 If ChrID = 22 Then CurrentLineText = CurrentLineText & Clipboard.GetText: DrawConsole: Exit Sub
 If ChrID = 16 Then
  Do
   cRevLn = cRevLn + 1
   tT = cConsoleText(UBound(cConsoleText) - cRevLn)
  Loop Until ((Left(tT, 1) <> Chr$(RET_LINE_CONST)) Or (cRevLn >= UBound(cConsoleText)))
  If (cRevLn >= UBound(cConsoleText)) Then Beep: cRevLn = 0: Exit Sub
  CurrentLineText = tT
  DrawConsole
  Exit Sub
 End If
 If ChrID = RET_LINE_CONST Then Exit Sub
 If ChrID = 8 Then
  If Not Len(CurrentLineText) = 0 Then CurrentLineText = Left(CurrentLineText, (Len(CurrentLineText) - 1))
  Call DrawConsole
  Exit Sub
 End If
 If ChrID = 13 Then
  cRevLn = 0
  Call AddLine(CurrentLineText)
  CurrentLineText = vbNullString
  Call ExecuteLine
 Else
  CurrentLineText = CurrentLineText & Chr$(ChrID)
 End If
 DoFlashEvent = True
 Call DrawConsole
End Sub

Public Sub DrawConsole()
Dim cY As Long, i As Long, cTl As Long, cR As RECT, cRa As RECT, cText As String
Dim bc As Long, fCn As Long, fCr As Long
 Call GetClientRect(hwnd, cR)
 Call GetCols(bc, fCn, fCr)
 client.Cls
 cTl = CreateBrush(bc)
 Call FillRect(hDC, cR, cTl)
 Call DeleteObject(cTl)
 Call Remap
 cY = Int(cR.Bottom / tHeight)
 cY = cY - 1
 cTl = (UBound(cConsoleText) - (cY - 1) - 1)
 If cTl < LBound(cConsoleText) Then cTl = LBound(cConsoleText)
 cY = -(tHeight)
 For i = cTl To UBound(cConsoleText)
  cY = cY + tHeight
  Call SetRect(cRa, (cR.Left + (tWidth)), (cY), cR.Right, (tHeight + cY))
  If Left(cConsoleText(i), 1) = Chr$(RET_LINE_CONST) Then
   cText = Right(cConsoleText(i), (Len(cConsoleText(i)) - 1))
   Call SetTextColor(hDC, fCr)
  Else
   cText = Prompt & cConsoleText(i)
   Call SetTextColor(hDC, fCn)
  End If
  Call DrawText(hDC, cText, Len(cText), cRa, DT_LEFT Or DT_VCENTER)
 Next i
 Call SetRect(cRa, (cR.Left + (tWidth)), (cY), cR.Right, (tHeight + cY))
 cText = Prompt & CurrentLineText
 Call DrawText(hDC, cText, Len(cText), cRa, DT_LEFT Or DT_VCENTER)
 If DoFlashEvent Then cTl = CreateBrush(vbBlack) Else cTl = CreateBrush(vbButtonFace)
 Call Remap(cText)
 i = tWidth
 Call Remap
 i = (tWidth) + i
 Call SetRect(cRa, (i), (cY), (i + tWidth), (tHeight + cY))
 Call FillRect(hDC, cRa, cTl)
 Call DeleteObject(cTl)
End Sub

Private Sub GetCols(ByRef backCol As Long, ByRef normFore As Long, ByRef respFore As Long)
Dim vN As String, dV, cS As String, bc As Long, nF As Long, rF As Long
 Call cVarC.ProcessLine("cvar_backCol", vN, dV, bc, cS)
 Call cVarC.ProcessLine("cvar_foreNorm", vN, dV, nF, cS)
 Call cVarC.ProcessLine("cvar_foreResp", vN, dV, rF, cS)
 backCol = bc
 normFore = nF
 respFore = rF
End Sub

Private Sub Class_Initialize()
 ReDim cConsoleText(0)
 ReDim cLineLocs(0)
End Sub

Private Sub Class_Terminate()
 Call ExitTimer(hwnd)
 Set client = Nothing
End Sub

Private Sub ExecuteLine()
Dim vN As String, dV As Variant, cV As Variant, cS As String, cRet As Long, rlc As String
 If LCase(cConsoleText(UBound(cConsoleText) - 1)) = "/quit" Then Unload client
 If cVarC Is Nothing Then Exit Sub
 rlc = Chr$(cCons.gRLC)
 cRet = cVarC.ProcessLine(cConsoleText(UBound(cConsoleText) - 1), vN, dV, cV, cS)
 Select Case cRet
  Case &H3234CDE: Call AddLine(rlc & """" & vN & """ is: current: """ & cV & """ default: """ & dV & """")
  Case &H2342AAA: Call AddLine(rlc & """" & vN & """ is: current: """ & cV & """ default: """ & dV & """")
  Case &HACCEDB:
  Case &H232AAB:  Call AddLine(rlc & """" & vN & """ is readonly")
  Case &H232DDE:  Call AddLine(rlc & """" & vN & """ is hidden")
  Case &H2EEC:    Call AddLine(rlc & """" & vN & """ cannot be read due to internal error")
  Case &H2322DDD: Call AddLine(rlc & """" & vN & """ not defined")
  Case &H21392:   Call AddLine(rlc & "the switch """ & cS & """ is not valid")
  Case &H21312BB: Call AddLine(rlc & """" & vN & """ cannot be used as a variable")
  Case &H1232AA:  Call AddLine(rlc & dV)
  Case &H9239B:   Call AddLine(rlc & "command executed")
  Case &H23CD:    Call AddLine(rlc & "command not executed")
  Case &H2398B:   Call AddLine(rlc & "the file """ & cV & """ is invalid")
 End Select
 DrawConsole
End Sub
