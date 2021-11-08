VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7350
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GET CACHED PASSWORDS"
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7350
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call GetPasswords
End Sub
