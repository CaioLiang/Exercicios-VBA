VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} exemplo02 
   Caption         =   "UserForm1"
   ClientHeight    =   2985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5145
   OleObjectBlob   =   "exemplo02.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "exemplo02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub okBtn_Click()
Dim A, B As Integer
A = Val(TextBox1)
B = Val(TextBox2)
If A > B Then
    lblResultado.Caption = Str(B) + ", " + Str(A)
End If
If B > A Then
    lblResultado.Caption = Str(A) + ", " + Str(B)
End If
If A = B Then
    lblResultado.Caption = Str(A) + " = " + Str(B)
End If
End Sub
