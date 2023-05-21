VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ex005 
   Caption         =   "Média Ponderada"
   ClientHeight    =   3210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5700
   OleObjectBlob   =   "ex005.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ex005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub calcularBtn_Click()
Dim n1, n2, n3, p1, p2, p3, media, totalPeso As Integer
n1 = Val(TextBox1)
n2 = Val(TextBox2)
n3 = Val(TextBox3)
p1 = Val(txtP1)
p2 = Val(txtP2)
p3 = Val(txtP3)
If p1 < 1 Or p2 < 1 Or p3 < 1 Then
    lblResultado.Caption = "Os pesos devem ter valores positivos!"
Else
    If n1 < 0 Or n1 > 10 Or n2 < 0 Or n2 > 10 Or n3 < 0 Or n3 > 10 Then
        lblResultado.Caption = "As notas devem variar de 0 a 10!"
    Else
        media = (n1 * p1 + n2 * p2 + n3 * p3) / (p1 + p2 + p3)
        If media < 5 Then
            lblResultado.Caption = "Sua média é: " + Str(media) + ", conceito E."
        Else
            If media < 6 Then
                lblResultado.Caption = "Sua média é: " + Str(media) + ", conceito D."
            Else
                If media < 7 Then
                    lblResultado.Caption = "Sua média é: " + Str(media) + ", conceito C."
                Else
                    If media < 8 Then
                        lblResultado.Caption = "Sua média é: " + Str(media) + ", conceito B."
                    Else
                        lblResultado.Caption = "Sua média é: " + Str(media) + ", conceito A."
                    End If
                End If
            End If
        End If
    End If
End If
End Sub
