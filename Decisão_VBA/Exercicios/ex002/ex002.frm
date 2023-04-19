VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ex002 
   Caption         =   "Maioridade"
   ClientHeight    =   2790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3390
   OleObjectBlob   =   "ex002.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ex002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub verificarBtn_Click()
Dim idade As Single
idade = Val(TextBox1)
If idade = 2023 Then
    lblResultado.Caption = "Jesus Cristo é você?"
Else:
    If idade >= 18 Then
        lblResultado.Caption = "Maior de idade."
    Else
        lblResultado.Caption = "Menor de idade."
    End If
End If
End Sub
