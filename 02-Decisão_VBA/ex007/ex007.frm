VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ex007 
   Caption         =   "Aumento de Salário"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3555
   OleObjectBlob   =   "ex007.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ex007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub verificarBtn_Click()
Dim cod As Integer
cod = Val(txtBoxCod)
If cod < 1 Or cod > 5 Then
    MsgBox ("Código Inválido!")
    lblCargo.Caption = ""
    lblPercentual.Caption = ""
Else
    If cod = 1 Then
        lblCargo.Caption = "Escrituário"
        lblPercentual.Caption = "50%"
    Else
        If cod = 2 Then
            lblCargo.Caption = "Secretária"
            lblPercentual.Caption = "35%"
        Else
            If cod = 3 Then
                lblCargo.Caption = "Caixa"
                lblPercentual.Caption = "20%"
            Else
                If cod = 4 Then
                    lblCargo.Caption = "Gerente"
                    lblPercentual.Caption = "10%"
                Else
                    lblCargo.Caption = "Diretor"
                    lblPercentual.Caption = "5%"
                End If
            End If
        End If
    End If
End If
End Sub
