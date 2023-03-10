VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "Conversor de Cº para Fahrenheit"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5775
   OleObjectBlob   =   "ex004.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCon_Click()
Dim grauC, grauF As Single
grauC = Val(txtC.Value)
grauF = (9 * grauC + 160) / 5
lblRes.Caption = Str(grauF)
End Sub
