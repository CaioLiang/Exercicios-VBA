VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Média Aritmética"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4215
   OleObjectBlob   =   "ex003.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnCal_Click()
Dim v1, v2, v3, v4, mF As Single
v1 = Val(txtA.Value)
v2 = Val(txtB.Value)
v3 = Val(txtC.Value)
v4 = Val(txtD.Value)
mF = (v1 + v2 + v3 + v4) / 4
lblRes.Caption = Str(mF)

End Sub
