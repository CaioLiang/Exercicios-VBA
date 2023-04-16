VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} exemplo04 
   Caption         =   "UserForm1"
   ClientHeight    =   3630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6585
   OleObjectBlob   =   "exemplo04.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "exemplo04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnMenor_Click()
Dim A, B As Integer
A = Val(TextBox1)
B = Val(TextBox2)
If A > B Then
    lblResultado.Caption = Str(B) + " é menor que " + Str(A)
Else
    If A = B Then
        lblResultado.Caption = Str(A) + " é igual a " + Str(B)
    Else
        lblResultado.Caption = Str(A) + " é menor que " + Str(B)
    End If
End If
End Sub

Private Sub crescenteBtn_Click()
Dim A, B As Integer
A = Val(TextBox1)
B = Val(TextBox2)
If A > B Then
    lblResultado.Caption = Str(B) + ", " + Str(A)
Else
    If A = B Then
        lblResultado.Caption = Str(A) + " é igual a " + Str(B)
    Else
        lblResultado.Caption = Str(A) + ", " + Str(B)
    End If
End If
End Sub

Private Sub decrescenteBtn_Click()
Dim A, B As Integer
A = Val(TextBox1)
B = Val(TextBox2)
If A > B Then
    lblResultado.Caption = Str(A) + ", " + Str(B)
Else
    If A = B Then
        lblResultado.Caption = Str(A) + " é igual a " + Str(B)
    Else
        lblResultado.Caption = Str(B) + ", " + Str(A)
    End If
End If
End Sub

Private Sub limparBtn_Click()
TextBox1.Text = ""
TextBox2.Text = ""
lblResultado.Caption = ""
End Sub

Private Sub maiorBtn_Click()
Dim A, B As Integer
A = Val(TextBox1)
B = Val(TextBox2)
If A > B Then
    lblResultado.Caption = Str(A) + " é maior que " + Str(B)
Else
    If A = B Then
        lblResultado.Caption = Str(A) + " é igual a " + Str(B)
    Else
        lblResultado.Caption = Str(B) + " é maior que " + Str(A)
    End If
End If
End Sub

Private Sub sairBtn_Click()
End
End Sub
