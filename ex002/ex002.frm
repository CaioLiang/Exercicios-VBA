VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ex002 
   Caption         =   "Mercadinho do Hiro"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5265
   OleObjectBlob   =   "ex002.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ex002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCal_Click()

Dim preP, qntP, total As Single
preP = Val(preProd.Value)
qntP = Val(qntProd.Value)
total = preP * qntP
lblResultado.Caption = Str(total)

End Sub
