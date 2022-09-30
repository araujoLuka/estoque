VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cadForm 
   Caption         =   "Formulario de Cadastro de Produto"
   ClientHeight    =   5040
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8124
   OleObjectBlob   =   "cadForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cadForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private pRng As Range
Private r As Boolean, p As Boolean
Private backup As Variant

' Procedimento ao iniciar o formulario
Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim cllr As String
    Dim rw As Integer
    
    Set ws = ActiveSheet
    Set tbl = ws.ListObjects(1)
    cllr = Application.Caller
    
    If (cllr Like "edit_*") Then
        rw = Right(cllr, Len(cllr) - InStr(1, cllr, "_"))
        Set pRng = tbl.ListRows(rw).Range
        Call preenchecadForm(pRng)
        cadFast = True
    End If
    
End Sub

' Botao que chama o procedimento de cadastro/atualizacao de produto
Private Sub cadBtn_Click()
    Dim vet() As Variant
    Dim i As Integer, n_box As Integer
    n_box = countFormTBX(Me)

    ' Impede a execucao se os campos obrigatorios nao estiverem preenchidos
    If (Not validaForm(Me, Me.Name, n_box)) Then Exit Sub
    
    Application.ScreenUpdating = False
    
    vet = geraVetorCad(Me, n_box)
    
    ' Somente atualiza o produto se ele ja estiver cadastrado
    If (cadCheck) Then
        Call atualizaProduto(vet(), pRng)
    Else 'Senao cadastra o produto
        Call cadastraProduto(vet)
    End If
    
    If (Not cadFast) Then
        Call resetForm 'Reseta o formulario depois de cadastrar/atualizar produto
    Else
        Unload Me
    End If

    Application.ScreenUpdating = True

End Sub

' Botao que chama o procedimento de remocao de produto
Private Sub remBtn_Click()

    ' Somente remove se estiver cadastrado
    If (cadCheck) Then
        ' Confirmacao de exclusao de produto
        If (MsgBox("Deseja realmente excluir o produto '" + box3 + "'?", vbYesNo) = vbYes) Then
            Call removeProduto(pRng) ' Sub de exclusao de produto
            Call resetForm ' Sub para resetar formulario
        End If
    End If

End Sub

' Botao que cancela a operacao de cadastro/atualizacao de produto
Private Sub cancelBtn_Click()

    ' Se atualizando um produto e nao eh cadastro rapido, reseta o formulario
    If (cadCheck And Not cadFast) Then
        Call resetForm
    Else ' Senao, fecha o formulario
        Unload Me
    End If

End Sub

Private Sub box1Check_Change()
    
    If (box1Check) Then
        Call travaCampo(box1, "SEM GTIN")
    Else
        Call destravaCampo(box1)
    End If

End Sub

Private Sub box1_Enter()
    If (box1 <> "") Then
        backup = box1
    End If
End Sub

' Procedimento ao preencher o codigo de barras no formulario
Private Sub box1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    If (r Or cadCheck) Then Exit Sub
    r = True
    
    If (box1 = "" Or UCase(box1) = "SEM GTIN") Then
        r = False
        Exit Sub
    End If
    ' Busca o produto na tabela por codigo de barras
    Set pRng = buscaProduto(box1, 1)
    
    ' Verifica se o produto foi encontrado
    ' Se encontrado, preenche o formulario com as informacoes cadastradas
    If (Not (pRng Is Nothing)) Then
        If (backup <> box1) Then
            Call preenchecadForm(pRng)
            p = True
        End If
    ' Se nao encontrado, verifica se houve alteracoes em informacoes preenchidas anteriormente
    ' > Exemplo[1] no fim do procedimento
    ElseIf (cadCheck) Then
        backup = box1 ' Salva o codigo de barras presente no formulario
        Call resetForm  ' Reseta o formulario
        box1 = backup ' Recupera o codigo salvo
    End If
    
    r = False

'------------------------------------------------------------------------------
'[1]
'> Exemplo {
'       Insere Em 'box1' um codigo ja cadastrado
'       Com isso sera preenchido o formulario com as informacoes do produto
'       e sai do procedimento
'
'       Em seguida, insere outro codigo em 'box1' mas nao cadastrado
'
'       Como o codigo de barras eh individual e exclusivo, reseta o formulario
'       mantendo o codigo inserido por ultimo para um novo cadastro
'}

End Sub

Private Sub box2_Enter()

    If (p) Then
        box4.SetFocus
    End If
    
End Sub

' Procedimento ao preencher o codigo interno no formulario
Private Sub box2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    If (r Or cadCheck) Then Exit Sub
    r = True
    
    If (box2 = "") Then
        r = False
        Exit Sub
    End If
    
    ' Busca o produto na tabela por codigo interno
    Set pRng = buscaProduto(box2, 2)

    ' Se encontrado, preenche o form com as informacoes cadastradas
    If (Not (pRng Is Nothing)) Then
        If (Not cadCheck) Then  ' Verifica se o formulario nao esta preenchido com outras informacoes cadastradas
            Call preenchecadForm(pRng)
            p = True
        End If
        ' Apos preenchido, verifica se o codigo do form eh diferente da tabela
        If (CLng(box2) <> pRng(1, 4)) Then
            box2.BackColor = RGB(255, 255, 0) ' Se diferente, preenche o campo na cor amarela
        Else
            box2.BackColor = vbWhite
        End If
    End If
    
    r = False
    
End Sub

Private Sub box3_enter()

    If (p) Then
        box4.SetFocus
        p = False
    End If
    
End Sub

' Procedimento ao preencher o nome do produto no formulario
Private Sub box3_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    ' Verifica se o formulario esta preenchido e se o nome do produto no form eh diferente da tabela
    If (cadCheck) Then
        If (box3 <> pRng.Cells(1, 5)) Then
            box3.BackColor = RGB(255, 255, 0) ' Se verdadeiro, preenche o campo na cor amarela
        End If
    End If
    
End Sub

Private Sub box4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    p = False
    
    ' Verifica se o formulario esta preenchido e se limite de estoque no form eh diferente da tabela
    If (cadCheck) Then
        If (CInt(box4) <> pRng.Cells(1, 6)) Then
            box4.BackColor = RGB(255, 255, 0) ' Se verdadeiro, preenche o campo na cor amarela
        End If
    End If
End Sub
