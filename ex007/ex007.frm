VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Média Ponderada"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7455
   OleObjectBlob   =   "ex007.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCal_Click()
Dim N1, N2, P1, P2, M As Single

N1 = Val(txtN1.Value)
N2 = Val(txtN2.Value)
P1 = Val(txtP1.Value)
P2 = 1 - P1
M = (N1 * P1) + (N2 * P2)

lblRes.Caption = "A média ponderada é: " & Str(M)
End Sub
