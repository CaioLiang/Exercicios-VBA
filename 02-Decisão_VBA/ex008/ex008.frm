VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ex008 
   Caption         =   "Seguros Takemasa"
   ClientHeight    =   5505
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
Dim idade As Integer, risco As String, nivel As Integer, verif As Boolean
idade = Val(txtBoxIdade)
risco = TextBox1.Value
If risco <> "a" And risco <> "b" And risco <> "m" Then
    MsgBox ("Os valores aceitos no grupo de risco são: a, m OU b!")
    lblResultado.Caption = ""
    verif = False
Else
    verif = True
End If
If verif Then
    Select Case idade
        Case 18 To 24
            Select Case risco
                Case "b": nivel = 7
                Case "m": nivel = 8
                Case Else: nivel = 9
            End Select
        Case 25 To 40
            Select Case risco
                Case "b": nivel = 4
                Case "m": nivel = 5
                Case Else: nivel = 6
            End Select
        Case 41 To 70
            Select Case risco
                Case "b": nivel = 1
                Case "m": nivel = 2
                Case Else: nivel = 3
            End Select
        Case Else
            MsgBox "Apenas pessoas entre 18 e 70 anos podem adquirir apólices."
            lblResultado.Caption = ""
    End Select
    If nivel <> 0 Then
        lblResultado.Caption = nivel
    End If
End If
End Sub

Private Sub Label3_Click()
MsgBox ("Digite a inicial do seu grupo de risco: b - baixo, m - medio, a - alto")
End Sub
