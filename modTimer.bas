Attribute VB_Name = "modTimer_Console"
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

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Public DoFlashEvent As Boolean
Private cTimerID As Long

Public Sub CreateTimer(ByVal hwnd As Long, ByVal cInterval As Long)
 cTimerID = SetTimer(hwnd, 0, cInterval, AddressOf TimerProc)
End Sub

Public Function ExitTimer(ByVal hwnd As Long) As Long
 ExitTimer = KillTimer(hwnd, cTimerID)
End Function

Public Sub TimerProc()
 DoFlashEvent = Not DoFlashEvent
 cCons.DrawConsole
End Sub
