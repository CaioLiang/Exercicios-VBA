VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} exemplo03 
   Caption         =   "UserForm1"
   ClientHeight    =   3180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "exemplo03.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Exemplo03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub executarBtn_Click()
Dim A As Integer
A = Val(TextBox1)
If A > 0 Then
    lblResultado.Caption = "O valor � positivo"
Else
    If A = 0 Then
        lblResultado.Caption = "O valor � nulo"
    Else
        lblResultado.Caption = "O valor � negativo"
    End If
End If
End Sub
