VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Info"
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo errH
Select Case LCase(Command) 'convert command to lowercase

Case "/help", "/?"
MsgBox "Help message", vbQuestion
Label1 = "Comand is " & Command

Case "/msg"
MsgBox "You have started the program with '/msg' command", vbInformation
Label1 = "Command is " & Command

'Case "/about", "/a", "-about", "-a" 'example to show about window
'code to show about window

'Case "/options"
'code to show options

'Case "-mycommand", "/mycmd" 'example myCommand
'code for myCommand

Case Else

If Command = "" Then Label1 = "No Command Given" Else Label1 = "I don't know this command '" & Command & "'"

End Select
Exit Sub
errH:
MsgBox Err.Description
End Sub
