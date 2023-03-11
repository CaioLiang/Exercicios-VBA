VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Calcular quantidade Salário Mín."
   ClientHeight    =   4230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4335
   OleObjectBlob   =   "ex009.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCal_Click()
Dim sal, salM, qntd As Single

sal = Val(txtSL.Value)
salM = Val(txtSM.Value)
qntd = sal / salM

lblRes.Caption = "Você ganha " & Str(qntd) & " vezes o salário mínimo."
End Sub

Private Sub duvida_Click()
MsgBox ("O salário mínimo atual (03/2023) é de 1.302 reais.")
End Sub
