VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Nick's Variable Console"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9390
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFC0C0&
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   406
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   626
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
 cCons.AddCharToLine KeyAscii
End Sub

Private Sub Form_Load()
Dim rlc As String
 Set cVarC = New clsVarCol
 Set cCons = New clsConsole
 Set cCons.Target = Me
 rlc = Chr$(cCons.gRLC)
 Call cCons.AddLine(rlc & "  *************************************************************************")
 Call cCons.AddLine(rlc & "  * Welcome to Nick's Variable Console app                                *")
 Call cCons.AddLine(rlc & "  * This is an API drawn console that is <relatively> bug free            *")
 Call cCons.AddLine(rlc & "  * The syntax is as follows:                                             *")
 Call cCons.AddLine(rlc & "  *  CREATING VARIABLE :: set <variable> <value> </readonly> </noview>    *")
 Call cCons.AddLine(rlc & "  *  READING VARIABLE :: <variable>                                       *")
 Call cCons.AddLine(rlc & "  *  UPDATING VARIABLE :: set <variable> <value>                          *")
 Call cCons.AddLine(rlc & "  *  RESETTING VARIABLE :: reset <variable>                               *")
 Call cCons.AddLine(rlc & "  * To update color scheme:                                               *")
 Call cCons.AddLine(rlc & "  *  BACKCOLOR :: set cvar_backCol <value>                                *")
 Call cCons.AddLine(rlc & "  *  FORECOLOR :: set cvar_foreNorm <value>                               *")
 Call cCons.AddLine(rlc & "  *  RESPONSECOLOR :: set cvar_foreResp <value>                           *")
 Call cCons.AddLine(rlc & "  * :NOTES:                                                               *")
 Call cCons.AddLine(rlc & "  *  that all variables are case sensitive                                *")
 Call cCons.AddLine(rlc & "  *  typing <Ctrl> + <p> will scroll through previously typed entries     *")
 Call cCons.AddLine(rlc & "  *  the char """ & rlc & """ is invalid                                              *")
 Call cCons.AddLine(rlc & "  *  typing ""/quit"" will exit the app                                     *")
 Call cCons.AddLine(rlc & "  *  only fixed size fonts are supported                                  *")
 Call cCons.AddLine(rlc & "  * I may be contacted about this app at any time at nebunagtar@yahoo.com *")
 Call cCons.AddLine(rlc & "  * UPDATES:                                                              *")
 Call cCons.AddLine(rlc & "  *  - added command and constant handling (June 4,2003)                  *")
 Call cCons.AddLine(rlc & "  *  - type '/emailauthor' to email me automatically                      *")
 Call cCons.AddLine(rlc & "  *  - now supports <Ctrl> + <v> (paste) command                          *")
 Call cCons.AddLine(rlc & "  *************************************************************************")
End Sub

Private Sub Form_Resize()
 Call cCons.Remap
 Call cCons.DrawConsole
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set cVarC = Nothing
 Set cCons = Nothing
 End
End Sub

