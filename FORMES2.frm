VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   2565
   LinkTopic       =   "Form2"
   ScaleHeight     =   3975
   ScaleWidth      =   2565
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   3825
      Left            =   0
      Picture         =   "FORMES2.frx":0000
      Top             =   120
      Width           =   2550
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image1_Click()
Unload Me
Load Form3
Form3.Show
End Sub
