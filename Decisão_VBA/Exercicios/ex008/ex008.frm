VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ex008 
   Caption         =   "Seguros Takemasa"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3750
   OleObjectBlob   =   "ex008.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ex008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub classificarBtn_Click()
Dim idade As Integer, risco As String, verif As Boolean
idade = Val(txtBoxIdade)
risco = TextBox1.Value
If risco <> "a" And risco <> "b" And risco <> "m" Then
    MsgBox ("Os valores aceitos no grupo de risco são: a, m OU b!")
    lblResultado.Caption = ""
    verif = False
Else
    verif = True
End If

If verif = True Then
    If idade < 18 Or idade > 70 Then
        MsgBox ("Apenas pessoas entre 18 e 70 anos podem adquirir apólices.")
        lblResultado.Caption = ""
    Else
        If idade < 25 Then
            If risco = "b" Then
                lblResultado.Caption = "7"
            Else
                If risco = "m" Then
                    lblResultado.Caption = "8"
                Else
                    lblResultado.Caption = "9"
                End If
            End If
        Else
            If idade < 41 Then
                If risco = "b" Then
                    lblResultado.Caption = "4"
                Else
                    If risco = "m" Then
                        lblResultado.Caption = "5"
                    Else
                        lblResultado.Caption = "6"
                    End If
                End If
            Else
                If idade < 71 Then
                    If risco = "b" Then
                        lblResultado.Caption = "1"
                    Else
                        If risco = "m" Then
                            lblResultado.Caption = "2"
                        Else
                            lblResultado.Caption = "3"
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub Label3_Click()
MsgBox ("Digite a inicial do seu grupo de risco: b - baixo, m - medio, a - alto")
End Sub
