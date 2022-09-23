Attribute VB_Name = "cadScripts"
' Autor: Lucas Araujo
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
        Call destravaCampo(.box5)
        .box5 = ""
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
    Dim rw As Integer, cl As Integer
    Dim arr As Variant
    
    rw = pRng.Row - pRng.ListObject.HeaderRowRange.Row + 1
    arr = Sheets("Estoque").ListObjects(1).Range.Value2
    For cl = 1 To UBound(arr, 2)
        If (arr(1, cl) = "ESTOQUE") Then Exit For
    Next
    
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
        Call travaCampo(.box5, arr(rw, cl))
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
    j = 1
    i = 1
    
    Set ws = Sheets("Cadastro")
    Set cTabble = ws.ListObjects(1)
    Set prodRow = insereRow(cTabble)
    Call editIcon_add(ws, prodRow, ws.ListObjects(1).ListRows.Count, 1)
    Call remIcon_add(ws, prodRow, ws.ListObjects(1).ListRows.Count, 2)

    prodRow.Formula = vet
    
    Call sortCad(cTabble)
    
    Set ws = Sheets("Estoque")
    Set cTabble = ws.ListObjects(1)
    Set prodRow = insereRow(cTabble)
       
    x = geraVetorEstoque(prodRow, vet(7), "", vet)
    
    prodRow.Formula = x
    With cTabble
        prodRow.ClearFormats
        .HeaderRowRange.Offset(1).Copy
        .DataBodyRange.PasteSpecial xlPasteFormats
    End With
    Application.CutCopyMode = False

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
    
    Call sortCad(pRow.ListObject)
    
    Set pRow = buscaProduto(vet(4), 2)
    pRow.Select
    
    MsgBox "Produto '" & pName & "' atualizado com sucesso!"

End Sub

Sub removeProduto(Optional pRow As Range)
    Dim ws As Worksheet
    Dim cTabble As ListObject
    Dim pCod As Integer, i As Integer, rw As Integer
    Dim arr As Variant
    Dim pNm As String, icon1 As String, icon2 As String
    
    rw = trataCaller(Application.Caller, icon1)
    If (rw = 0) Then
        rw = pRow.Row - pRow.ListObject.HeaderRowRange.Row
        icon1 = "rem_" & rw
        icon2 = "edit_" & rw
    Else
        icon2 = "edit_" & rw
    End If
    
    pNm = pRow(1, 5)
    pCod = pRow(1, 4)
    
    Set ws = Sheets("Estoque")
    Set cTabble = ws.ListObjects(1)
    arr = cTabble.DataBodyRange.Value2
    
    For i = 1 To cTabble.ListRows.Count
        If (pCod = arr(i, 3)) Then
            Call deleteRow(ws, cTabble.ListRows(i).Range)
            Exit For
        End If
    Next

    Set ws = Sheets("Cadastro")

    Call ajustaIcon(ws, rw, icon1)
    Call ajustaIcon(ws, rw, icon2)
    Call deleteRow(ws, pRow)
    
    Call sortCad(cTabble)

    MsgBox "Produto '" & pNm & "' removido com sucesso!"

End Sub

Function geraVetorCad(u As UserForm, ByVal n_box As Integer) As Variant
    Dim i As Integer, j As Integer
    Dim vet(1 To 7) As Variant
    Dim rng As Range
    Dim ws As Worksheet
    Set rng = ActiveSheet.ListObjects(1).HeaderRowRange.Offset(1)
    i = 1
    j = 1
    
    vet(j) = "=COUNTA(" & rng(1, j + 1).Address & ":[@[CODIGO DE BARRAS]])"
    j = j + 1
    
    vet(j) = u.Controls("box" & i)
    j = j + 1
    
    vet(j) = "=IF([@[CODIGO INTERNO]]="""","""",IF([@[CODIGO INTERNO]]<1000,""AP"",""P�""))"
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
        
    geraVetorCad = vet
End Function