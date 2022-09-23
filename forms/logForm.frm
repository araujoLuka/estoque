VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} logForm 
   Caption         =   "Autenticação de Acesso"
   ClientHeight    =   2640
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3888
   OleObjectBlob   =   "logForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "logForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub logginBtn_Click()
    Dim key As String
    Dim tp As Integer
    
    If (user = "close") Then Application.Quit
    
    key = buscaAcesso(user, tp)
    
    If (passw <> key) Then
        Call invalidPass
    Else
        Call planAccess(user, tp)
    End If

End Sub

Private Sub logginBtn_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    Call anotherPass

End Sub

Private Sub cancelBtn_Click()
    
    Unload Me
    
End Sub

Private Sub Label3_Click()
    Dim us As String
    Dim key As String
    Dim tp As Integer
    
    Me.Hide
    us = InputBox("Insira seu nome de usuario:", "Recuperação de Acesso")
    key = buscaAcesso(us, tp)
    
    If (key = "-1") Then
        MsgBox "Usuario não encontrado!" & vbCrLf & _
                    "Entre em contato com o administrador", vbExclamation
    Else
        MsgBox "Voce pode redefinir depois!"
    End If
    
End Sub
