VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ex004 
   Caption         =   "Média"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4215
   OleObjectBlob   =   "ex004.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ex004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub calcularBtn_Click()
Dim n1, n2, n3, media As Integer
n1 = Val(TextBox1)
n2 = Val(TextBox2)
n3 = Val(TextBox3)
media = (n1 + n2 + n3) / 3
If n1 < 0 Or n1 > 10 Or n2 < 0 Or n2 > 10 Or n3 < 0 Or n3 > 10 Then
    lblResultado.Caption = "As notas devem variar de 0 a 10!"
Else
    If media < 5 Then
        lblResultado.Caption = "reprovado"
    Else
        If media < 7 Then
            lblResultado.Caption = "exame"
        Else
            lblResultado.Caption = "aprovado"
        End If
    End If
End If
End Sub
