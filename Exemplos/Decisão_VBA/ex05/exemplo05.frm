VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} exemplo05 
   Caption         =   "UserForm1"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3750
   OleObjectBlob   =   "exemplo05.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "exemplo05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub calcularBtn_Click()
Dim A, B, X As Single
A = Val(TextBox1)
B = Val(TextBox2)
If A = 0 Then
    lblResultado.Caption = "O Coeficiente A deve ser diferente de 0!"
Else
    X = -B / A
    lblResultado.Caption = "O valor de X é: " + Str(X)
End If
End Sub
