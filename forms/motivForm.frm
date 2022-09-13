VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} motivForm 
   Caption         =   "Motivo da Movimentação do Estoque"
   ClientHeight    =   3744
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9396
   OleObjectBlob   =   "motivForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "motivForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancelBtn_Click()
    
    bCancel = True
    Me.Hide
End Sub

Private Sub sendBtn_Click()
    
    Me.Hide
End Sub

Private Sub opt_o_Change()
    
    If (opt_o) Then
        Call destravaCampo(opt_o_txt)
    Else
        Call travaCampo(opt_o_txt)
    End If
End Sub

