VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVarCol"
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

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Type VARIABLE_STRUCT
 cVarName  As String
 cDefVal   As Variant
 cCurVal   As Variant
 cReadOnly As Boolean
 cViewable As Boolean
End Type

Public Enum cRetCodes
 VAR_UPDATED = &H3234CDE
 VAR_OK = &H2342AAA
 VAR_CREATED = &HACCEDB
 VAR_READONLY = &H232AAB
 VAR_HIDDEN = &H232DDE
 VAR_NOT_UPDATABLE = &H2EEC
 VAR_NOT_PRESENT = &H2322DDD
 VAR_UNKNOWN_SWITCH = &H21392
 VAR_NOT_CREATABLE = &H21312BB
 KW_RETURNED = &H1232AA
 COM_OK = &H9239B
 COM_ERR = &H23CD
 COM_BADFILE = &H2398B
End Enum

Private VarStruct() As VARIABLE_STRUCT

Public Function GetRetCodeText(ByVal cRetCode As cRetCodes) As String
 Select Case cRetCode
  Case cRetCodes.VAR_CREATED:           GetRetCodeText = "var created"
  Case cRetCodes.VAR_HIDDEN:            GetRetCodeText = "var hidden"
  Case cRetCodes.VAR_NOT_PRESENT:       GetRetCodeText = "var not defined"
  Case cRetCodes.VAR_NOT_UPDATABLE:     GetRetCodeText = "syntax error"
  Case cRetCodes.VAR_OK:                GetRetCodeText = "var success"
  Case cRetCodes.VAR_READONLY:          GetRetCodeText = "var readonly"
  Case cRetCodes.VAR_UPDATED:           GetRetCodeText = "var updated"
  Case cRetCodes.VAR_UNKNOWN_SWITCH:    GetRetCodeText = "unknown switch specified"
  Case cRetCodes.VAR_NOT_CREATABLE:     GetRetCodeText = "word reserved"
  Case cRetCodes.KW_RETURNED:           GetRetCodeText = "keyword val returned"
  Case cRetCodes.COM_OK:                GetRetCodeText = "command executed"
  Case cRetCodes.COM_ERR:               GetRetCodeText = "command not executed"
  Case cRetCodes.COM_BADFILE:           GetRetCodeText = "target file not found"
  Case Else:                            GetRetCodeText = "unknown ret code"
 End Select
End Function

Public Function ProcessLine(ByVal cLine As String, Optional ByRef vN As String, Optional ByRef dV As Variant, Optional ByRef cV As Variant, Optional ByRef cS As String) As cRetCodes
' :: LINE_SYNTAX ::
'    Setting Var:   Set <varname> <value> </readonly> </noview>
'    Getting Var:   <varname>
'    Update Var :   Set <varname> <value>
'    Reset Var  :   Reset <varname>
' :: NOTES ::
'    Setting Var:   the /readonly and /noview switch can only be applied to var setting
'                   [switch]    [allow var modify]   [allow var view]
'                   /readonly :         no         :        yes
'                   /noview   :         yes        :        no
'    Getting Var:   var will not return if the cViewable flag is False
'    Update Var :   var will not update if the cReadOnly flag is True
'    Reset Var  :   resets var's cVal (current val) to dVal (default val)
' :: SYNTAX RULES ::
'    For Commands:  all commands are proceeded by the '/' char (ie. /shellexecute)
'                   commands are not case sensitive
'                   variable commands ('set','reset') are not considered part of the command structure
'    For Constant:  all constants are proceeded by the '%' char (ie. %date)
'                   constants may be used for variable values (ie. 'set d %date')
Dim plStr() As String, i As Long
Dim tRO As Boolean, tNV As Boolean, tVN As Variant, tVal As Variant, cBL As Boolean
 If Trim(cLine) = vbNullString Then Exit Function
 cLine = Trim(cLine)
 plStr = Split(cLine, " ")
 Select Case LCase(plStr(LBound(plStr)))
  Case "set"
   tVN = plStr(LBound(plStr) + 1)
   vN = tVN
   If tVN = "reset" Then ProcessLine = VAR_NOT_CREATABLE: Exit Function
   If LBound(plStr) + 2 > UBound(plStr) Then ProcessLine = VAR_NOT_UPDATABLE: Exit Function
   tVal = plStr(LBound(plStr) + 2)
   tRO = False
   tNV = True
   If GetIdx(tVN) = -1 Then
    If Not LBound(plStr) + 3 > UBound(plStr) Then
     Select Case LCase(plStr(LBound(plStr) + 3))
      Case "/readonly": tRO = True
      Case "/noview": tNV = False
      Case Else: cS = plStr(LBound(plStr) + 3): ProcessLine = VAR_UNKNOWN_SWITCH: Exit Function
     End Select
     If Not LBound(plStr) + 4 > UBound(plStr) Then
      Select Case LCase(plStr(LBound(plStr) + 4))
       Case "/readonly": tRO = True
       Case "/noview": tNV = False
       Case Else: cS = plStr(LBound(plStr) + 4): ProcessLine = VAR_UNKNOWN_SWITCH: Exit Function
      End Select
     End If
    End If
   End If
   dV = CheckKeyword(tVal)
   ProcessLine = CreateVar(tVN, dV, cV, tRO, tNV)
  Case "reset"
   tVN = plStr(LBound(plStr) + 1)
   vN = tVN
   i = GetIdx(tVN)
   If Not i = -1 Then
    VarStruct(i).cCurVal = VarStruct(i).cDefVal
    cV = VarStruct(i).cCurVal
    dV = VarStruct(i).cDefVal
    ProcessLine = VAR_OK
   Else
    ProcessLine = VAR_NOT_PRESENT
   End If
  Case Else
   tVN = plStr(LBound(plStr))
   vN = CheckKeyword(tVN)
   If Not (vN = tVN) Then
    dV = vN
    ProcessLine = KW_RETURNED
    Exit Function
   End If
   vN = CheckCommand(tVN, cBL)
   If Not (vN = tVN) Then
    dV = vN
    If cBL Then
     If (LBound(plStr) + 1 <= UBound(plStr)) Then
      cV = plStr(LBound(plStr) + 1)
      ProcessLine = ExecuteCommand(vN, plStr(LBound(plStr) + 1))
     Else
      ProcessLine = COM_BADFILE
     End If
     Exit Function
    Else
     ProcessLine = ExecuteCommand(vN, vbNullString)
     Exit Function
    End If
   End If
   ProcessLine = GetVar(tVN, dV, cV)
 End Select
End Function

Private Function ExecuteCommand(ByVal cCommand As String, ByVal lParam As Variant) As cRetCodes
 Select Case cCommand
  Case "/ShellExecuteA"
   If Not FileExists(CStr(lParam)) Then ExecuteCommand = COM_BADFILE: Exit Function
   If ShellExecute(cCons.hwnd, "open", CStr(lParam), vbNullString, vbNullString, 1) = 42 Then ExecuteCommand = COM_OK Else ExecuteCommand = COM_ERR
  Case "/EmailAuthorA"
   If ShellExecute(cCons.hwnd, "open", "mailto:nebunagtar@yahoo.com?subject=Feedback on Console", vbNullString, vbNullString, 1) = 42 Then ExecuteCommand = COM_OK Else ExecuteCommand = COM_ERR
  Case "/OpenURLA"
   If ShellExecute(cCons.hwnd, "open", CStr(lParam), vbNullString, vbNullString, 1) = 42 Then ExecuteCommand = COM_OK Else ExecuteCommand = COM_ERR
  Case Else: ExecuteCommand = COM_ERR
 End Select
End Function

Private Function CheckCommand(ByVal cCommand As String, ByRef reqlParam As Boolean) As String
 Select Case LCase(cCommand)
  Case "/shellexecute": CheckCommand = "/ShellExecuteA": reqlParam = True
  Case "/emailauthor":  CheckCommand = "/EmailAuthorA":  reqlParam = False
  Case "/openurl":      CheckCommand = "/OpenURLA":      reqlParam = True
  Case Else:            CheckCommand = cCommand
 End Select
End Function

Private Function CheckKeyword(ByVal cWord As String) As String
 Select Case LCase(cWord)
  Case "%date":         CheckKeyword = Date
  Case "%time":         CheckKeyword = Time
  Case "%now":          CheckKeyword = Now
  Case "%path":         CheckKeyword = App.Path
  Case "%version":      CheckKeyword = App.Major & "." & App.Minor & "." & App.Revision
  Case "%compiledate":  CheckKeyword = "June 4,2003"
  Case Else:            CheckKeyword = cWord
 End Select
End Function

Public Function CreateVar(ByVal vN As String, ByRef dV As Variant, Optional ByRef cV As Variant = vbNullString, Optional ByVal ro As Boolean = False, Optional ByVal v As Boolean = True) As cRetCodes
 If Not (IsMissing(cV)) Then If (cV = vbNullString) Then cV = dV
 If Not GetIdx(vN) = -1 Then CreateVar = UpdateVar(vN, cV, cV, dV): Exit Function
 ReDim Preserve VarStruct(LBound(VarStruct) To (UBound(VarStruct) + 1))
 With VarStruct(UBound(VarStruct))
  .cVarName = vN
  .cDefVal = dV
  .cCurVal = cV
  .cReadOnly = ro
  .cViewable = v
 End With
 CreateVar = VAR_CREATED
End Function

Public Function UpdateVar(ByVal vN As String, Optional ByVal cVn As Variant = vbNullString, Optional ByRef cV As Variant, Optional ByRef dV As Variant) As cRetCodes
Dim i As Long
 i = GetIdx(vN)
 If Not i = -1 Then
  If VarStruct(i).cReadOnly Then UpdateVar = VAR_READONLY: Exit Function
  If cVn = vbNullString Then cVn = VarStruct(i).cDefVal
  VarStruct(i).cCurVal = cVn
  dV = VarStruct(i).cDefVal
  cV = VarStruct(i).cCurVal
  UpdateVar = VAR_UPDATED
 Else
  UpdateVar = VAR_NOT_PRESENT
 End If
End Function

Public Function GetVar(ByVal vN As String, Optional ByRef dV As Variant, Optional ByRef cV As Variant, Optional ByRef ro As Boolean) As cRetCodes
Dim i As Long
 i = GetIdx(vN)
 If Not i = -1 Then
  If Not VarStruct(i).cViewable Then GetVar = VAR_HIDDEN: Exit Function
  With VarStruct(i)
   dV = .cDefVal
   cV = .cCurVal
   ro = .cReadOnly
  End With
  GetVar = VAR_OK
 Else
  GetVar = VAR_NOT_PRESENT
 End If
End Function

Private Function GetIdx(ByVal vN As String) As Long
Dim i As Long
 If UBound(VarStruct) = LBound(VarStruct) Then GetIdx = -1: Exit Function
 For i = LBound(VarStruct) To UBound(VarStruct)
  If VarStruct(i).cVarName = vN Then GetIdx = i: Exit Function
 Next i
 GetIdx = -1
End Function

Private Function FileExists(FileName As String) As Boolean
 FileExists = Dir(FileName) <> ""
End Function

Private Sub Class_Initialize()
 ReDim VarStruct(0)
End Sub

Private Sub Class_Terminate()
 Erase VarStruct
 ReDim VarStruct(0)
End Sub
