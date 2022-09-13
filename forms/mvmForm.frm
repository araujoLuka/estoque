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

Private Sub regButton_Click()
    Dim vet As Variant
    Dim mtv As String
    Dim i As Integer, n_box As Integer
    n_box = countFormTBX(Me)
    
    mtv = defineMotiv(Me)
    If (IsEmpty(mtv)) Then Exit Sub
      
    For i = 0 To pList.ListCount - 1
        vet = geraVetorMov(Me, Me.Name, pList.List(i, 0), mtv, n_box, i)
        If (IsEmpty(vet)) Then Exit Sub
        If (vet(8) > 0) Then
            Call regEntrada(vet)
        Else
            Call regSaida(vet)
        End If
        Call regMovimentacao(vet)
        Call atualizaEstoque(buscaProduto(2, vet(6)), vet(8))
    Next
    
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim cols_wdth As String
    cols_wdth = "70;80;70;100;10"

    With hList
        .ColumnWidths = cols_wdth
        .AddItem
        .List(0, 0) = "COD. HERD"
        .List(0, 1) = "COD. BARRAS"
        .List(0, 2) = "COD. INT."
        .List(0, 3) = "PRODUTO"
        .List(0, 4) = "QTD"
    End With
    pList.ColumnWidths = cols_wdth
    
    boxU = Range("actv")
    boxData = Date
    boxHora = Time
    
End Sub

Private Sub addBtn_Click()
    Dim index As Integer
    Dim herd As Boolean
    Dim cod As String, cdInt As String, cdBar As String, cdHerd As String
    Dim pArray() As Variant
    
    cod = trataCodigo(box1, index, herd)
    If (herd) Then
        cdHerd = box1
    Else
        cdHerd = "SEM CH"
    End If
    
    pArray = buscaProduto(index, cod).Value2
    
    If (IsEmpty(pArray)) Then
        MsgBox "Impossivel movimentar produto não cadastrado"
        Exit Sub
    End If
        
    Call insereDadoLista(Me.pList, geraVetorMvm(cdHerd, pArray, box3))

    Call clearForm(Me)
    
End Sub

Private Sub subBtn_Click()
    Dim i As Integer, index As Integer
    Dim herd As Boolean
    Dim cod As String, cdInt As String, cdBar As String, cdHerd As String
    Dim pArray() As Variant
    i = pList.ListCount
    
    box3 = -box3
    cod = trataCodigo(box1, index, herd)
    If (herd) Then
        cdHerd = box1
    Else
        cdHerd = "SEM CH"
    End If
    
    pArray = buscaProduto(index, cod).Value2
    
    If (IsEmpty(pArray)) Then
        MsgBox "Impossivel movimentar produto não cadastrado"
        Exit Sub
    ElseIf (pArray(1, 6) < CInt(box3)) Then
        MsgBox "Impossivel remover mais de " & _
                pArray(1, 6) & " unidades" & vbCrLf & _
                "ERRO: Remoção maior que estoque atual"
        Exit Sub
    End If
    
    Call insereDadoLista(Me.pList, geraVetorMvm(cdHerd, pArray, box3))
    
    Call clearForm(Me)

End Sub

Private Sub box1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
    Dim cod As String
    Dim rng As Range
    Dim pArray As Variant
    
    cod = trataCodigo(box1, i)
    If (cod = "") Then Exit Sub
    
    Set rng = buscaProduto(i, cod)
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

