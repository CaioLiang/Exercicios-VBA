VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} exemplo01 
   Caption         =   "UserForm1"
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4770
   OleObjectBlob   =   "exemplo01.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "exemplo01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub executarBtn_Click()
Dim A As Single
A = Val(TextBox1.Value)

If A > 0 Then
    lblResultado.Caption = "O valor � positivo"
End If
If A = 0 Then
    lblResultado.Caption = "O valor � nulo"
End If
If A < 0 Then
    lblResultado.Caption = "O valor � negativo"
End If
End Sub
