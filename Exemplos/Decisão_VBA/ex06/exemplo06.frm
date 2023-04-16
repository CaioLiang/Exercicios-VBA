VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} exemplo06 
   Caption         =   "UserForm1"
   ClientHeight    =   4185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3960
   OleObjectBlob   =   "exemplo06.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "exemplo06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub calcularBtn_Click()
Dim A, B, C, X1, X2 As Integer
A = Val(TextBox1)
B = Val(TextBox2)
C = Val(TextBox3)
If A = 0 Then
    lblResultado.Caption = "O coeficiente A deve ser diferente de 0!"
Else:
    Delta = B * B - 4 * A * C
    If Delta > 0 Then
        X1 = (-B + Sqr(Delta)) / 2 * A
        X2 = (-B - Sqr(Delta)) / 2 * A
        lblResultado.Caption = "As duas raízes são: " + Str(X1) + ", " + Str(X2)
    Else
        If Delta = 0 Then
            X1 = (-B + Sqr(Delta)) / 2 * A
            lblResultado.Caption = "A única existente é: " + Str(X1)
        Else
            lblResultado.Caption = "Não existem raízes reais nesta função."
        End If
    End If
End If
End Sub
