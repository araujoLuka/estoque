VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} mvmForm 
   Caption         =   "Movimentação Por Lote"
   ClientHeight    =   7020
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13920
   OleObjectBlob   =   "mvmForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "mvmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    Dim cols_wdth As String
    cols_wdth = "60;70;60;130;10"

    With hList
        .ColumnWidths = cols_wdth
        .AddItem
        .List(0, 0) = "COD HERD"
        .List(0, 1) = "COD BARRAS"
        .List(0, 2) = "COD INT"
        .List(0, 3) = "PRODUTO"
        .List(0, 4) = "QTD"
    End With
    pList.ColumnWidths = cols_wdth
    
    boxU = Range("actv")
    boxData = Date
    boxHora = Time
    
End Sub

Private Sub regButton_Click()
    Dim vet As Variant
    Dim mtv_e As String, mtv_s As String
    Dim i As Integer, n_box As Integer, list_c As Integer
    n_box = countFormTBX(Me)
    list_c = pList.ListCount - 1
    
    If (list_c < 0) Then
        MsgBox "Sem lancamentos para registrar!", vbInformation
        Exit Sub
    End If
    
    If (Not defineMotivMult(pList.List, list_c, mtv_e, mtv_s)) Then Exit Sub
    
    For i = 0 To list_c
        If (pList.List(i, 4) > 0) Then
            vet = geraVetorMov(Me, Me.Name, pList.List(i, 0), mtv_e, n_box, i)
            Call regEntrada(vet)
        Else
            vet = geraVetorMov(Me, Me.Name, pList.List(i, 0), mtv_s, n_box, i)
            Call regSaida(vet)
        End If
        Call regMovimentacao(vet)
        Call atualizaEstoque(vet(6), vet(8))
    Next
    
    Unload Me
End Sub

Private Sub addBtn_Click()
    Dim index As Integer
    Dim herd As Boolean
    Dim cod As String, cdHerd As String
    Dim pArray() As Variant
    Dim n_box As Integer
    
    cod = trataCodigo(box1, index, herd)
    If (herd) Then
        cdHerd = box1
        ' Impede a execucao se houver duplicidade de codigo herdeiro
        If (Not validaMovim(cdHerd, 1)) Then
            Exit Sub
        End If
    Else
        cdHerd = "SEM CH"
    End If
    
    n_box = countFormTBX(Me)
    
    If (Not validaForm(Me, Me.Name, n_box)) Then Exit Sub
        
    pArray = buscaProduto(cod, index).Value2
    
    If (IsEmpty(pArray)) Then
        MsgBox "Impossivel movimentar produto não cadastrado"
        Exit Sub
    End If

    If (insereDadoLista(Me.pList, geraVetorMvm(cdHerd, pArray, box3), 0)) Then
        Call clearForm(Me)
    End If
    
End Sub

Private Sub subBtn_Click()
    Dim index As Integer
    Dim herd As Boolean
    Dim cod As String, cdHerd As String
    Dim pArray() As Variant
    Dim n_box As Integer
    
    cod = trataCodigo(box1, index, herd)
    If (herd) Then
        cdHerd = box1
        ' Impede a execucao se houver duplicidade de codigo herdeiro
        If (Not validaMovim(cdHerd, 1)) Then
            Exit Sub
        End If
    Else
        cdHerd = "SEM CH"
    End If
    
    n_box = countFormTBX(Me)
    
    If (Not validaForm(Me, Me.Name, n_box)) Then Exit Sub
    
    pArray = buscaProduto(cod, index).Value2
    
    If (IsEmpty(pArray)) Then
        MsgBox "Impossivel movimentar produto não cadastrado"
        Exit Sub
    End If
    
    ' Impede a insercao na lista se a quantidade a ser removida eh maior do que o estoque
    If (Not validaEstoque(box3, getEstoque(cod))) Then Exit Sub

    If (insereDadoLista(Me.pList, geraVetorMvm(cdHerd, pArray, -box3), 0)) Then
        Call clearForm(Me)
    End If

End Sub

Private Sub box1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
    Dim cod As String
    Dim rng As Range
    Dim pArray As Variant
    
    cod = trataCodigo(box1, i)
    If (cod = "") Then Exit Sub
    
    Set rng = buscaProduto(cod, i)
    If (rng Is Nothing) Then
        MsgBox "Codigo não cadastrado"
        Cancel = True
        box1.SetFocus
        Exit Sub
    End If
    
    pArray = rng.Value2
    
    box2 = pArray(1, 5)
        
End Sub

Private Sub box2_Enter()
    
    If (box1 <> "" And box2 <> "") Then
        Call travaCampo(box2, box2)
        box3.SetFocus
    Else
        Call destravaCampo(box2)
    End If
    
End Sub

Private Sub Label7_Click()
    
    bnmForm.Show

End Sub

Private Sub pList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    pList.RemoveItem (pList.ListIndex)

End Sub

