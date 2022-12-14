Attribute VB_Name = "xmlScripts"
Option Explicit

Sub entradaXML()
    
    Call movXML
    
End Sub

Sub movXML()
    Dim rng As Range
    Dim mat As Variant, arr As Variant
    Dim i As Integer, j As Integer
    Dim dt As Date, tm As Date
    Dim res As VbMsgBoxResult
    
    Application.ScreenUpdating = False
    
    dt = Date
    tm = Time
    
    Do
    mat = import_XML
    
    If (IsEmpty(mat)) Then Exit Sub
    
    With Sheets("Entrada").ListObjects(1)
        If (.ListRows.Count = 0) Then Exit Do
        If (Not .ListColumns(.ListColumns.Count).DataBodyRange.Find(("*" & mat(0))) Is Nothing) Then
            MsgBox "Nota " & mat(0) & " ja registrada!", vbCritical
        Else
            Exit Do
        End If
    End With
    Loop
    
    Load xmlForm
    For i = 1 To UBound(mat)
        Call insereDadoLista(xmlForm.pList, mat(i), 1)
    Next
    xmlForm.Show
    
    If (Not IsUserFormLoaded("xmlForm")) Then Exit Sub
    
    For i = 1 To UBound(mat, 1)
        Set rng = buscaProduto(mat(i)(2), 2)
        If (Not rng Is Nothing) Then
            arr = rng.Value2
            If (Not checkData(arr, mat(i))) Then
                arr = geraVetorCadXML(mat(i), arr(1, 6), getEstoque(arr(1, 4)))
                Call atualizaProduto(arr, rng)
            End If
        Else
            res = MsgBox(mat(i)(3) & " n?o cadastrado!" & vbCrLf & vbCrLf & _
                         "Deseja cadastrar?", vbQuestion + vbYesNo)
            If (res = vbYes) Then
                arr = geraVetorCadXML(mat(i))
                Call cadastraProduto(arr)
            End If
        End If
        arr = geraVetorMovXML(mat(i), mat(0), dt, tm)
        
        Call regEntrada(arr)
        Call regMovimentacao(arr)
        
        Call atualizaEstoque(mat(i)(2), mat(i)(4))
    Next
    
    Unload xmlForm
    Application.ScreenUpdating = True

End Sub

Function import_XML() As Variant
    Dim wb As Workbook
    Dim i As Integer, j As Integer, k As Integer
    Dim nf As Long
    Dim xArray As Variant, xTam As Integer
    Dim rng As Range, rng2 As Range
    Dim cols(0 To 4) As Variant, vet(1 To 4) As Variant
    Dim xMatrix As Variant
    
    k = 1
    i = 1
    
    Set wb = openXML()
    If (wb Is Nothing) Then
        Application.ScreenUpdating = True
        Exit Function
    End If
    xArray = wb.Worksheets(1).ListObjects(1).Range.Value2
    xTam = UBound(xArray, 2)
    
    nf = getData(xArray, xTam, i, "nNF", 0)
    cols(0) = getData(xArray, xTam, i, "Item", 1)
    cols(2) = getData(xArray, xTam, i, "cProd", 1)
    cols(1) = getData(xArray, xTam, i, "cEAN", 1)
    cols(3) = getData(xArray, xTam, i, "xProd", 1)
    cols(4) = getData(xArray, xTam, i, "qCom", 1)
    
    ReDim xMatrix(UBound(cols(0), 1)) As Variant
    
    xTam = UBound(xArray, 1)
    xMatrix(k - 1) = nf
    For i = 1 To xTam - 1
        If (IsNumeric(cols(0)(i)) And cols(0)(i) >= k) Then
            For j = 1 To UBound(vet)
                vet(j) = cols(j)(i)
            Next
            xMatrix(k) = vet
            k = k + 1
        End If
    Next
    
    ReDim Preserve xMatrix(k - 1)
    wb.Close False
    
    import_XML = xMatrix

End Function

Function openXML() As Workbook
    Dim XMLFile As String
        
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    XMLFile = Application.GetOpenFilename("Arquivo XML (*.xml), *.xml", MultiSelect:=False)
    If (XMLFile <> CStr(False)) Then
        Set openXML = Workbooks.openXML(Filename:=XMLFile, LoadOption:=xlXmlLoadImportToList)
    End If
    
    Application.DisplayAlerts = True
    
End Function

Function record_XML_data(ws As Worksheet, vet() As Variant, nf As Long) As Range
    Dim tbl As ListObject
    Dim tRow As Range
    Dim just_update As Boolean
    Dim i As Integer
    
    just_update = False
    Set tbl = ws.ListObjects(1)
    Set tRow = define_row(tbl, vet(2), nf, just_update)
    
    If (just_update = True) Then
        tRow(1, 5) = tRow(1, 5) + vet(UBound(vet))
    Else
        tRow = Split(nf & "," & Join(vet, ","), ",")
    End If
    Set record_XML_data = tRow
    
End Function

Function define_row(ByRef tbl As ListObject, ByVal cod As Integer, ByVal nf As Long, ByRef upd As Boolean) As Range
    Dim i As Integer
    Dim arr As Variant
    Set define_row = Nothing
    arr = tbl.DataBodyRange.Value2
    
    For i = 1 To tbl.ListRows.Count
        If (arr(i, 1) = "") Then
            Set define_row = tbl.ListRows(i).Range
            Exit For
        ElseIf (arr(i, 3) = cod And arr(i, 1) = nf) Then
            Set define_row = tbl.ListRows(i).Range
            upd = True
            Exit For
        End If
    Next
    If (define_row Is Nothing) Then
        Set define_row = tbl.ListRows.Add().Range
        tbl.ListRows(1).Range.EntireRow.Copy
        define_row.EntireRow.PasteSpecial xlPasteFormats
        define_row.EntireRow.Hidden = False
        Application.CutCopyMode = False
    End If
    
End Function

Function getData(arr As Variant, ByVal tam As Integer, ByRef i As Integer, _
                 ByVal crit As String, tp As Integer) As Variant
    Dim vet As Variant
    Dim j As Integer, t As Integer
    t = UBound(arr, 1) - 1
    ReDim vet(1 To t) As Variant
    
    For i = i To tam
        If (arr(1, i) Like ("*" & crit)) Then
            Select Case tp
                Case 0
                    getData = arr(2, i)
                Case 1
                    For j = 1 To t
                        vet(j) = arr(j + 1, i)
                    Next
                    getData = vet
            End Select
            i = i + 1
            Exit For
        End If
    Next
End Function

Function checkData(cad As Variant, vet As Variant) As Boolean
    Dim i As Integer
    Dim dif As Boolean
    Dim res As VbMsgBoxResult
    
    checkData = False
    
    For i = 1 To UBound(vet) - 1
        If (IsNumeric(vet(i)) Or i = 1) Then
            If (CStr(vet(i)) <> CStr(cad(1, i * 2))) Then
                Debug.Print vet(i) & " <> " & cad(1, i * 2)
                res = MsgBox("Existem dados desatualizados para o produto:" & vbCrLf & _
                             "'" & vet(3) & "'" & vbCrLf & vbCrLf & _
                             "CODIGO DO CADASTRO: " & cad(1, i * 2) & vbCrLf & _
                             "CODIGO NA NOTA FISCAL: " & vet(i) & vbCrLf & vbCrLf & _
                             "Deseja atualizar?", vbQuestion + vbYesNo)
                If (res = vbNo) Then
                    vet(i) = cad(1, i * 2)
                Else
                    dif = True
                End If
            End If
        Else
            If (vet(i) <> cad(1, i + 2)) Then
                Debug.Print vet(i) & " <> " & cad(1, i + 2)
                res = MsgBox("Existem dados desatualizados para o produto:" & vbCrLf & _
                             "'" & vet(3) & "'" & vbCrLf & vbCrLf & _
                             "NOME DO CADASTRO: " & cad(1, i + 2) & vbCrLf & _
                             "NOME NA NOTA FISCAL: " & vet(i) & vbCrLf & vbCrLf & _
                             "Deseja atualizar?", vbQuestion + vbYesNo)
                If (res = vbNo) Then
                    vet(i) = cad(1, i + 2)
                Else
                    dif = True
                End If
            End If
        End If
    Next
    
    checkData = True
    
End Function
