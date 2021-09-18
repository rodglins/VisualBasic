VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5805
   LinkTopic       =   "Form3"
   ScaleHeight     =   4470
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   2295
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   3135
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Load Form1
Form1.Show
End Sub
