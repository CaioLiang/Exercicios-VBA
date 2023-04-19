VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ex001 
   Caption         =   "Média Aritmética"
   ClientHeight    =   4350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3945
   OleObjectBlob   =   "ex001.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ex001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub calcularBtn_Click()
Dim A, B, C, D, media As Single
A = Val(TextBox1)
B = Val(TextBox2)
C = Val(TextBox3)
D = Val(TextBox4)
If A < 0 Or A > 10 Or B < 0 Or B > 10 Or C < 0 Or C > 10 Or D < 0 Or D > 10 Then
    lblResultado.Caption = "As notas devem variar de 0 a 10!"
Else
    media = (A + B + C + D) / 4
    If media >= 7 Then
        lblResultado.Caption = "Sua média é: " + Str(media) + ", você foi APROVADO!"
    Else
        lblResultado.Caption = "Sua média é: " + Str(media) + ", você foi REPROVADO!"
    End If
End If
End Sub
