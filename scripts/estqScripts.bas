Attribute VB_Name = "estqScripts"
Option Explicit

Sub criaEstoque()
    Dim ws As Worksheet
    Dim t1 As ListObject, t2 As ListObject, t3 As ListObject
    Dim lRow As ListRow
    Dim rng As Range
    Dim x As Variant
    Dim estq As Integer, i As Integer
    Dim obsv As String
    i = 1
    
    Set ws = Sheets("Estoque")
    Set t1 = Sheets("Cadastro").ListObjects(1)
    Set t2 = ws.ListObjects(1)
    Set t3 = Workbooks("Estoque JC (2)").Sheets("Cadastro").ListObjects(1)
    
    For Each lRow In t1.ListRows
        Set rng = insereRow(t2)
        estq = t3.ListRows(i).Range(1, 6)
        obsv = t3.ListRows(i).Range(1, 7)
        i = i + 1
        x = geraVetorEstoque(lRow.Range, estq, obsv)
        rng.Formula = x
        With t2
            rng.ClearFormats
            .HeaderRowRange.Offset(1).Copy
            .DataBodyRange.PasteSpecial xlPasteFormats
        End With
        Application.CutCopyMode = False
    Next
    
End Sub

Sub atualizaEstoque(ByVal cod As Integer, ByVal qtd As Variant)
    Dim ws As Worksheet
    Dim eRng As Range
    Dim i As Integer
    
    Set ws = Sheets("Estoque")
    Set eRng = buscaProduto(cod, 2, ws)
    
    If (eRng Is Nothing) Then Exit Sub
    
    qtd = CInt(qtd)
    
    For i = 1 To eRng.Count
        If (eRng.ListObject.HeaderRowRange(1, i) = "ESTOQUE") Then
            Exit For
        End If
    Next
    
    eRng.Cells(1, i) = eRng.Cells(1, i) + qtd
End Sub

Sub removeEstoque(ByVal pCod As Integer)
    Dim ws As Worksheet
    Dim cTabble As ListObject
    Dim arr As Variant
    Dim i As Integer

    Set ws = Sheets("Estoque")
    Set cTabble = ws.ListObjects(1)
    
    arr = cTabble.ListColumns(4).Range.Value2
    
    For i = 1 To UBound(arr, 1)
        If (pCod = arr(i, 1)) Then
            Call deleteRow(ws, cTabble.ListRows(i - 1).Range)
            Exit For
        End If
    Next
End Sub

Function getEstoque(ByVal cod As Integer) As Integer
    Dim rng As Range
    
    getEstoque = -1
    
    Set rng = buscaProduto(cod, 2, Sheets("Estoque"))
    If (rng Is Nothing) Then Exit Function
    
    getEstoque = rng.Value2(1, 7)

End Function

Function geraVetorEstoque(pRow As Range, Optional ByVal estq As Integer, _
                            Optional ByVal obsv As String, Optional ByVal vetCad As Variant) As Variant
    Dim vet(1 To 9) As Variant
    Dim i As Integer
    Dim tName As String
    
    tName = Sheets("Cadastro").ListObjects(1).Name
    
    If (IsMissing(vetCad)) Then
        For i = 1 To 5
            vet(i) = pRow(1, i).Formula
        Next
        vetCad = vet
    End If
    
    For i = 1 To 5
        vet(i) = vetCad(i)
    Next
    vet(i) = "=INDEX(" & tName & "[LIMITE DE ESTOQUE]," & _
                "MATCH([@[CODIGO INTERNO]]," & tName & "[CODIGO INTERNO],0))"
    i = i + 1
    vet(i) = estq
    i = i + 1
    vet(i) = obsv
    i = i + 1
    vet(i) = "=IF($H" & pRow.Row & "<$G" & pRow.Row & ",""COMPRAR URGENTE""," & _
                "IF($H" & pRow.Row & "<=CEILING.MATH($G" & pRow.Row & "*1.6),""ESTOQUE BAIXO"",""OK""))"

    geraVetorEstoque = vet
End Function

Sub criaLista(ByVal lst As Variant, ByVal tam As Integer, ByVal tp As Integer)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim i As Integer
    Dim fName As String
    
    Application.ScreenUpdating = False
    
    Set ws = Sheets("Lista")
    Set tbl = ws.ListObjects(1)
    
    For i = 0 To tam - 1
        Set rng = insereRow(tbl)
        rng = lst(i)
    Next
    
    ws.PageSetup.PrintArea = ws.Range("B1", tbl.Range).Address
    
    Select Case tp
        Case 1
            tbl.HeaderRowRange = Array("CODIGO", "PRODUTO", "COMPRA MINIMA", "COMPRA RECOMENDADA")
            ws.Range("B1") = "LISTA DE COMPRA (" & Format(Date, "dd/mm/yyyy") & ")"
            fName = "Lista de Compra"
        
        Case 2
            tbl.HeaderRowRange = Array("CODIGO", "PRODUTO", "ESTOQUE ATUAL", "ESTOQUE LIMITE")
            ws.Range("B1") = "LISTA DE ESTOQUE BAIXO (" & Format(Date, "dd/mm/yyyy") & ")"
            fName = "Lista de Estoque Baixo"
    
    End Select
    
    ws.ExportAsFixedFormat xlTypePDF, fName, 1, OpenAfterPublish:=True
    
    tbl.DataBodyRange.Delete
    ws.Range("B1", tbl.HeaderRowRange) = ""
    
    Application.ScreenUpdating = True

End Sub

Sub listaCompra()
    Dim arr As Variant, lst As Variant, aux(3) As Variant
    Dim i As Integer, tam As Integer
    tam = 0
    
    arr = Sheets("Estoque").ListObjects(1).DataBodyRange.Value2
    
    ReDim lst(UBound(arr, 1))
    
    For i = 1 To UBound(arr, 1)
        If (arr(i, UBound(arr, 2)) Like "COMPRA*") Then
            aux(0) = arr(i, 3)
            aux(1) = arr(i, 4)
            aux(2) = arr(i, 5) - arr(i, 6)
            aux(3) = WorksheetFunction.Ceiling_Math(arr(i, 5) * 1.2) - arr(i, 6)
            lst(tam) = aux
            tam = tam + 1
        End If
    Next
    
    If (tam = 0) Then
        MsgBox "Não há produtos para compra urgente!", vbExclamation
        Exit Sub
    End If
    
    Call criaLista(lst, tam, 1)
    
End Sub

Sub listaEstoqueB()
    Dim arr As Variant, lst As Variant, aux(3) As Variant
    Dim i As Integer, tam As Integer
    tam = 0
    
    arr = Sheets("Estoque").ListObjects(1).DataBodyRange.Value2
    
    ReDim lst(UBound(arr, 1))
    
    For i = 1 To UBound(arr, 1)
        If (arr(i, UBound(arr, 2)) Like "ESTOQUE*") Then
            aux(0) = arr(i, 3)
            aux(1) = arr(i, 4)
            aux(2) = arr(i, 6)
            aux(3) = arr(i, 5)
            lst(tam) = aux
            tam = tam + 1
        End If
    Next
    
    If (tam = 0) Then
        MsgBox "Não há produtos com estoque baixo!", vbExclamation
        Exit Sub
    End If
    
    Call criaLista(lst, tam, 2)
    
End Sub

