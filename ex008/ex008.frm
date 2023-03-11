VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Banco do Hiro"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4875
   OleObjectBlob   =   "ex008.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnTaxa_Click()
Dim taxaJ, deposito, rend, rendTotal As Single

taxaJ = Val(txtTJ.Value)
deposito = Val(txtVD.Value)
rend = (deposito * taxaJ) / 100
rendTotal = deposito + rend

lblRes.Caption = "O rendimento é: " & Str(rend) & "R$. Já o rendimento total é: " & Str(rendTotal) & "R$."
End Sub
