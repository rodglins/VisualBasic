VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000012&
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   3600
   ClientTop       =   2805
   ClientWidth     =   8745
   FillColor       =   &H000000FF&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   8745
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   8655
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000002&
         Caption         =   "Clique aqui"
         Height          =   375
         Index           =   1
         Left            =   6840
         MaskColor       =   &H000000FF&
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Text            =   "CLIQUE NO BOTÃO AO LADO - - -  - - -           - - -  - - - -    - - --- -  -   -   - -  - - ---  -   -  - - >"
         Top             =   240
         Width           =   6495
      End
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

For cont = 1 To 2

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

If cont = 2 Then
MsgBox ("A:") & a
MsgBox ("B:") & B
MsgBox ("C:") & C
MsgBox ("D:") & D
MsgBox ("E:") & E
MsgBox ("F:") & F
End If

Next
End Sub
