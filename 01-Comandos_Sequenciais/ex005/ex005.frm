VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Calcular volume de um Barril d'óleo"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ex005.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCal_Click()
Dim Altu, Volu, Raio As Single
Altu = Val(txtA.Value)
Raio = Val(txtR.Value)
Volu = 3.14159 * (Raio * Raio) * Altu
lblRes.Caption = Str(Volu)
End Sub
