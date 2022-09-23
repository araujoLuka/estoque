VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} movForm 
   Caption         =   "Formulario de Movimentacao de Estoque"
   ClientHeight    =   6972
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8796
   OleObjectBlob   =   "movForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "movForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pRng As Range

Private Sub UserForm_Initialize()
    Dim entry As String, cod As String
    Dim index As Integer
    Dim herd As Boolean
    Dim mbResult As VbMsgBoxResult
    index = 2
    herd = False
    
    boxU = Range("actv")
    
    entry = InputBox("Leia ou digite o codigo do produto")
    If (entry = "") Then
        Unload Me
        Exit Sub
    End If
    
    cod = trataCodigo(entry, index, herd)
    If (cod = "") Then
        MsgBox "Codigo informado invalido!", vbCritical
        Unload Me
        Exit Sub
    End If

    Set pRng = buscaProduto(cod, index)
    
    If (pRng Is Nothing) Then
        mbResult = MsgBox("Produto nao encontrado na base de dados" & vbCrLf & _
               "Deseja cadastrar o produto?", vbQuestion + vbYesNo)
        If (mbResult = vbYes) Then
            Load cadForm
            With cadForm
                .Controls("box" & index) = cod
                .Controls("box" & index + 1).SetFocus
                .cadFast = True
                .Show
            End With
            Set pRng = buscaProduto(cod, index)
            If (pRng Is Nothing) Then
                MsgBox "Falha ao cadastrar produto!"
            End If
        End If
    End If
    Call preencheMovForm(pRng, index, herd, entry, cod)
    If (ActiveSheet.Name = "Entrada") Then
        subBtn.Visible = False
    ElseIf (ActiveSheet.Name = "Saida") Then
        addBtn.Visible = False
        subBtn.Left = addBtn.Left
    End If
    
End Sub

' Botao para adicionar produto no estoque
Private Sub addBtn_Click()
    Dim vet() As Variant
    Dim mtv As String
    Dim n_box As Integer
    
    mtv = trataMotiv(Me)
    n_box = countFormTBX(Me)

    ' Impede a execucao se houver falha na motivacao da movimentacao
    If (Not validaMotiv(mtv)) Then Exit Sub

    ' Impede a execucao se os campos obrigatorios nao estiverem preenchidos
    If (Not validaForm(Me, Me.Name, n_box)) Then Exit Sub
    
    ' Impede a execucao se houver duplicidade de codigo herdeiro
    If (Not validaMovim(boxH, 1)) Then Exit Sub
    
    vet = geraVetorMov(Me, Me.Name, boxH, mtv, n_box)
    
    Application.ScreenUpdating = False
    
    Call regEntrada(vet())
    Call regMovimentacao(vet())
    Call atualizaEstoque(vet(6), box4)
    
    Application.ScreenUpdating = True
    
    Unload Me
    On Error Resume Next
    movForm.Show
End Sub

' Botao para subtrair produto do estoque
Private Sub subBtn_Click()
    Dim vet() As Variant
    Dim mtv As String
    Dim n_box As Integer
    
    mtv = trataMotiv(Me)
    n_box = countFormTBX(Me)

    ' Impede a execucao se houver falha na motivacao da movimentacao
    If (Not validaMotiv(mtv)) Then Exit Sub

    ' Impede a execucao se os campos obrigatorios nao estiverem preenchidos
    If (Not validaForm(Me, Me.Name, n_box)) Then Exit Sub
    
    ' Impede a execucao se houver duplicidade de codigo herdeiro
    If (Not validaMovim(boxH, 1)) Then Exit Sub
    
    ' Impede a execucao se a quantidade a ser removida eh maior do que o estoque
    If (Not validaEstoque(box4, buscaProduto(box2, 2, Sheets("Estoque"))(1, 6))) Then Exit Sub
    
    box4 = -box4
    vet = geraVetorMov(Me, Me.Name, boxH, mtv, n_box)
    Call regSaida(vet())
    Call regMovimentacao(vet())
    Call atualizaEstoque(vet(6), box4)
    
    Unload Me
    On Error Resume Next
    movForm.Show
End Sub

Private Sub opt_o_Change()
    If (opt_o) Then
        Call destravaCampo(opt_o_txt)
    Else
        Call travaCampo(opt_o_txt)
    End If
End Sub

' Botao que cancela a operacao de cadastro/atualizacao de produto
Private Sub cancelBtn_Click()

    Unload Me

End Sub

Private Sub usrBtn_Click()
    
    Call loggin
    boxU = Range("actv")
    
End Sub
