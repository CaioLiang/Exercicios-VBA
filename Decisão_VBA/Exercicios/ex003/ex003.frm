VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ex003 
   Caption         =   "Menor número"
   ClientHeight    =   2790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4200
   OleObjectBlob   =   "ex003.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ex003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub menorBtn_Click()
Dim n1, n2 As Integer
n1 = Val(TextBox1)
n2 = Val(TextBox2)
If n1 = n2 Then
    lblResultado.Caption = "Os valores são iguais."
Else
    If n1 < n2 Then
        lblResultado.Caption = "O menor número é: " + Str(n1)
    Else
        lblResultado.Caption = "O menor número é: " + Str(n2)
    End If
End If
End Sub
