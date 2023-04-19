VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ex010 
   Caption         =   "Gasto de Energia"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4425
   OleObjectBlob   =   "ex010.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ex010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCal_Click()
Dim salM, quiloW, cadaQW, totalP, novoVal As Single

salM = Val(txtSM.Value)
quiloW = Val(txtQW.Value)
cadaQW = (salM / 5)
totalP = quiloW * (salM / 5)
novoVal = totalP * 0.85

lblA.Caption = "O valor em reais de cada quilowatt é: " & Str(cadaQW) & "R$."
lblB.Caption = "O valor em reais a ser pago é: " & Str(totalP) & "R$."
lblC.Caption = "O novo valor a ser pago, com um desconto de 15% é: " & Str(novoVal) & "R$."

End Sub

Private Sub duvida_Click()
MsgBox ("O salário mínimo atual (03/2023) é de 1.302 reais.")
End Sub

