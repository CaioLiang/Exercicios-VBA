VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Produto de 4 valores"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4365
   OleObjectBlob   =   "ex001.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnProd_Click()

Dim v1, v2, v3, v4, prod As Integer

v1 = Val(txt1.Value)
v2 = Val(txt2.Value)
v3 = Val(txt3.Value)
v4 = Val(txt4.Value)
prod = v1 * v2 * v3 * v4
lblResultado.Caption = Str(prod)

End Sub
