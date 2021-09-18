VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ver Nota "
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Nota:"
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Nota 3"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Pesquisa Escolar do Ano Letivo de 2000"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "Nota 2"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Nota 1"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nome"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim nome As String
Dim nota1 As Double
Dim nota2 As Double
Dim nota3 As Double
Dim media As Double

nome = Text1.Text
nota1 = Text2.Text
nota2 = Text3.Text
nota3 = Text4.Text

media = (nota1 + nota2 + nota3) / 3

Label7.Caption = media

End Sub
