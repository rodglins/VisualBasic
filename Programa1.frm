VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000012&
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   3600
   ClientTop       =   2805
   ClientWidth     =   4680
   FillColor       =   &H000000FF&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   4680
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   480
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   360
      Picture         =   "Programa1.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   5475
      TabIndex        =   1
      Top             =   1440
      Width           =   5535
      Begin VB.HScrollBar HScroll1 
         Height          =   1575
         Left            =   1080
         TabIndex        =   2
         Top             =   120
         Width           =   30
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000002&
      Caption         =   "Clique aqui"
      Height          =   975
      Left            =   2400
      MaskColor       =   &H000000FF&
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "NANANANNNNNNNNNNNNNNNNNNOOOOOOOOOOOLL"
      Height          =   735
      Left            =   840
      TabIndex        =   3
      Top             =   5400
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   2400
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim L As String
a = 0
B = 0
C = 0
D = 0
E = 0
F = 0
cont = 1

For cont = 1 To 60

L = InputBox("Digite A LETRA")

If L = "A" Then
a = a + 1
ElseIf L = "B" Then
B = B + 1
ElseIf L = "C" Then
C = C + 1
ElseIf L = "D" Then
D = D + 1
ElseIf L = "E" Then
E = E + 1
ElseIf L = "F" Then
F = F + 1
End If

If cont = 60 Then
MsgBox ("A:") & a
MsgBox ("B:") & B
MsgBox ("C:") & C
MsgBox ("D:") & D
MsgBox ("E:") & E
MsgBox ("F:") & F
End If

Next
End Sub
