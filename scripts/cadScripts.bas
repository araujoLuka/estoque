Attribute VB_Name = "cadScripts"
' Autor: Lucas Araujo
' Ultima atualizacao: 09/07/2022
'
' Modulo para procedimentos do formulario de cadastro/atualizacao de produto, 'cadForm'

Option Explicit

' Inicializa o formulario para cadastrar produto
Sub iniciaCadastro()

    cadForm.Show

End Sub

' Inicializa o formulario para atualizar produto
Sub iniciaAtualiz()

    cadForm.Show

End Sub

' Resetar o formulario
Sub resetForm()
    Dim i As Integer
   
    With cadForm
        For i = 1 To countFormTBX(cadForm)
            .Controls("box" & i) = ""
            .Controls("box" & i).BackColor = vbWhite
        Next
        .cadCheck = False
        .cadCheck.Visible = False
        .cadBtn.Caption = "Cadastrar"
        .remBtn.Visible = False
        .box1Check = False
        .box1.SetFocus
    End With

End Sub

' Preenche o formulario com informacoes da tabela
' pRng eh a linha da tabela aonde estao as informacoes
Sub preenchecadForm(pRng As Range)
    Dim i As Integer
    
    With cadForm
        .Caption = Replace(.Caption, "Cadastro", "Atualização")
        If (pRng(1, 2) = "SEM GTIN") Then
            .box1Check = True
        Else
            .box1 = pRng(1, 2)
        End If
        .Controls("box" & 2) = pRng(1, 4)
        .Controls("box" & 3) = pRng(1, 5)
        .Controls("box" & 4) = pRng(1, 6)
        Call travaCampo(.boxE, "-")
        .cadCheck = True
        .cadCheck.Visible = True
        .cadBtn.Caption = "Atualizar"
        .remBtn.Visible = True
    End With

End Sub

Sub cadastraProduto(vet() As Variant)
    Dim ws As Worksheet
    Dim cTabble As ListObject
    Dim prodRow As Range
    Dim i As Integer, j As Integer
    Dim x As Variant
    Dim pName As String
    Set ws = Sheets("Cadastro")
    Set cTabble = ws.ListObjects(1)
    j = 1
    i = 1
    
    Set prodRow = insereRow(cTabble)
    
    prodRow = vet
    
    Call sortCad(cTabble)
        
    MsgBox "Produto '" & vet(5) & "' cadastrado com sucesso!"

End Sub

Sub atualizaProduto(vet As Variant, pRow As Range)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim pName As String, hdr As String
    Dim wsPaths As Variant, arr As Variant
    Dim i As Integer, w As Integer
    Dim c As Integer
    
    wsPaths = Array("Estoque")
    
    For i = 1 To UBound(vet) - 2
        If (Not vet(i) Like "=*") Then
            If (vet(i) <> CStr(pRow(1, i))) Then
                hdr = pRow.ListObject.HeaderRowRange(1, i)
                For w = 0 To UBound(wsPaths)
                    Set ws = Worksheets(wsPaths(w))
                    Set tbl = ws.ListObjects(1)
                    arr = tbl.HeaderRowRange.Value2
                    
                    For c = 1 To UBound(arr, 2)
                        If (arr(1, c) = hdr) Then Exit For
                    Next
                    
                    If (c > UBound(arr, 2)) Then Exit For
                                            
                    ws.ListObjects(1).ListColumns(c).DataBodyRange.Replace pRow(1, i), vet(i)
                Next
            End If
        End If
    Next
    
    pName = vet(5)
    pRow = vet
    
    MsgBox "Produto '" & pName & "' atualizado com sucesso!"

End Sub

Sub removeProduto(pRow As Range, nm As String)
    Dim ws As Worksheet
    Dim cTabble As ListObject
    Set ws = Sheets("Cadastro")
    Set cTabble = ws.ListObjects(1)
    
    Call deleteRow(ws, pRow, 0)
        
    Call sortCad(cTabble)
                       
    MsgBox "Produto '" & nm & "' removido com sucesso!"

End Sub

Function geraVetorCad(u As UserForm, ByVal n_box As Integer) As Variant
    Dim i As Integer, j As Integer
    Dim vet(1 To 7) As Variant
    Dim rng As Range
    Set rng = ActiveSheet.ListObjects(1).ListRows(1).Range
    i = 1
    j = 1
    
    vet(j) = "=COUNTA(" & rng(1, j + 1).Address & ":[@[CODIGO DE BARRAS]])"
    j = j + 1
    
    vet(j) = u.Controls("box" & i)
    j = j + 1
    
    vet(j) = "=IF([@[CODIGO INTERNO]]="""","""",IF([@[CODIGO INTERNO]]<1000,""AP"",""PÇ""))"
    j = j + 1
    
    With u
        For i = 2 To n_box
            vet(j) = .Controls("box" & i)
            If (Not IsNumeric(vet(j))) Then
                vet(j) = UCase(vet(j))
            End If
            j = j + 1
        Next
    End With
    
    vet(j) = ""
    
    geraVetorCad = vet
End Function
