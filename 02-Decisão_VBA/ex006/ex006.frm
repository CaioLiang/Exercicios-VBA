VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ex006 
   Caption         =   "Classifica��o de Idade"
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3540
   OleObjectBlob   =   "ex006.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ex006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub classificarBtn_Click()
Dim idade As Integer
idade = Val(txtBoxIdade)
If idade < 0 Then
    lblResultado.Caption = "A idade deve ser um n�mero maior ou igual a 0!"
Else
    If idade = 2023 Then
        lblResultado.Caption = "Jesus Cristo � voc�?"
    Else
        If idade < 3 Then
            lblResultado.Caption = "Rec�m-nascido"
        Else
            If idade < 12 Then
                lblResultado.Caption = "Crian�a"
            Else
                If idade < 20 Then
                    lblResultado.Caption = "Adolescente"
                Else
                    If idade < 61 Then
                        lblResultado.Caption = "Adulto"
                    Else
                        lblResultado.Caption = "Idoso"
                    End If
                End If
            End If
        End If
    End If
End If
End Sub
