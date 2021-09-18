VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Height          =   495
      Left            =   4680
      TabIndex        =   14
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1320
      TabIndex        =   13
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Top             =   2400
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Index           =   0
      Left            =   1560
      TabIndex        =   5
      Top             =   2040
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1560
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label Label7 
      Caption         =   "altura"
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "peso"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "idade"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "cor cabelo"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "cor olhos"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Sexo"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Nome"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim nm As String
Dim sx As String
Dim co As String
Dim cc As String
Dim id As Integer
Dim ps As Integer
Dim alt As Double
a = 0
b = 0
c = 0
d = 0
e = 0
f = 0
g = 0
p1 = 0
p2 = 0
p3 = 0
mm = 0
nome = 0
nome1 = 0
nome2 = 0
olhos = 0
cor = 0
soma = 0
soma2 = 0
med1 = 0
med2 = 0
ml = 999
peso = 0
ma = 0
cont = cont + 1
For cont = 1 To 5

nm = Text1.Text
sx = Text2.Text
co = Text3.Text
cc = Text4.Text
id = Text5.Text
ps = Text6.Text
alt = Text7.Text

If sx = fem Then
a = a + 1
ElseIf id > 15 Then
c = c + 1
p1 = c / a * 100
ElseIf id > mm And cc = cast Then
mm = id
nome = nm
olhos = co
End If

If sex = masc Then
b = b + 1
p2 = (b / cont) * 100
ElseIf co = cast Then
d = d + 1
End If

If co = ver And cc = cast Then
e = e + 1
soma = soma + id
med1 = soma / e
End If

If co = azuis And cc = preto Then
f = f + 1
p3 = (f / cont) * 100
End If

If ps < ml Then
nome1 = nm
sexo = sx
End If

If co = cast And cab = cast And id > 20 Then
g = g + 1
soma2 = soma2 + alt
med2 = soma2 / g
End If

If alt > malt Then
ma = alt
nome2 = nm
peso2 = ps
cor = cc
End If
Next
msg box = "a: " & a
msg box = "b: " & p2
msg box = "c: " & p1
msg box = "d: " & d
msg box = "e: " & med1
msg box = "f: " & p3
msg box = "g: " & med2
msg box = "h: " & nome1 & sexo
msg box = "i: " & ma & nome2 & peso & cor
msg box = "j: " & mm & nome & olhos

End Sub
