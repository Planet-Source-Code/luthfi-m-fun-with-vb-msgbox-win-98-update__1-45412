VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Custom MsgBox"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   180
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   870
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Custom Icon"
      Height          =   435
      Left            =   150
      TabIndex        =   2
      Top             =   1860
      Width           =   1365
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Custom Caption"
      Height          =   435
      Left            =   1650
      TabIndex        =   1
      Top             =   1860
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "No Button"
      Height          =   435
      Left            =   3210
      TabIndex        =   0
      Top             =   1860
      Width           =   1245
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'sub classing VBA.MsgBox function
'Call it the same way, just like regular msgbox call but with additional params
'VB always call the module level function declaration first before the class level function
'So our custom msgbox handler will be called first
  
Private Sub Command1_Click()
  MsgBox "Hello you all" & vbCr & "I Have no buttons..!!", vbYesNoCancel + vbInformation, , , , , , True
End Sub

Private Sub Command2_Click()
  Dim sCaptions As Variant
  sCaptions = Array("Yeah Right..", "Maybe..", "No Way")
  MsgBox "Hello you all", vbYesNoCancel + vbInformation, , , , sCaptions
End Sub

Private Sub Command3_Click()
  MsgBox "Hello you all", vbYesNoCancel + vbInformation, , , , , Me.Picture1
End Sub
