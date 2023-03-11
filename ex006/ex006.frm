VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Viagem da Boa"
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4890
   OleObjectBlob   =   "ex006.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnGerar_Click()
Dim velM, temG, litU, dist As Single
velM = Val(txtV.Value)
temG = Val(txtT.Value)
dist = velM * temG
litU = dist / 12
lblEst.Caption = "A velocidade média era de: " & Str(velM) & " km/h e o tempo gasto foi de: " & Str(temG) & " hora(s). Um total de: " & Str(dist) & " kms foram percorridos, gastando um total de: " & Str(litU) & " litro(s)."
End Sub
