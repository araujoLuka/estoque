Attribute VB_Name = "movScripts"
' Autor: Lucas Araujo
'
'
' Modulo para procedimentos do formulario de movimentacao de estoque, 'movForm'

Option Explicit

' Inicializa a movimentacao de estoque
Sub iniciaMovimentacao()
    
    movForm.Show

End Sub

Sub multiMoviment()
    
    mvmForm.Show

End Sub

' Preenche o formulario com informacoes da tabela
' pRng eh a linha da tabela aonde estao as informacoes
Sub preencheMovForm(pRng As Range, index As Integer, _
                    herd As Boolean, entry As String, cod As String)
    Dim i As Integer
    
    With movForm
        .boxData = Date
        .boxHora = Time
        If (herd) Then
            .boxH = entry
            Call travaCampo(.box4, 1)
            .opt_1.SetFocus
        Else
            If (.box4 = "") Then
                .box4 = 1
                .box4.SetFocus
            Else
                .opt_1.SetFocus
            End If
        End If
        
        If (pRng Is Nothing) Then
            MsgBox "ATENCAO!" & vbCrLf & _
                "Movimentacao de produto nao cadastrado" & vbCrLf & _
                "Controle de estoque pode estar comprometido!", vbExclamation
            .tempCheck = True
            .Controls("box" & index) = cod
            .Controls("box" & (index Mod 2) + 1) = "SEM CODIGO"
            If (.tempCheck) Then .box3 = "PRODUTO TEMPORARIO"
            .box3.SetFocus
        Else
            .Controls("box" & 1) = pRng.Cells(1, 2)
            .Controls("box" & 2) = pRng.Cells(1, 4)
            .Controls("box" & 3) = pRng.Cells(1, 5)
        End If
    End With
    
End Sub

Sub clearForm(u As UserForm)
    Dim i As Integer
    
    For i = 1 To countFormTBX(u)
        u.Controls("box" & i) = ""
        Call destravaCampo(u.Controls("box" & i))
    Next
    u.box1.SetFocus
End Sub

Sub regEntrada(vet As Variant)
    Dim cTabble As ListObject
    Dim prodRow As Range
    
    Set cTabble = Sheets("Entrada").ListObjects(1)
    Set prodRow = insereRow(cTabble)
    
    prodRow = vet
End Sub

Sub regSaida(vet As Variant)
    Dim cTabble As ListObject
    Dim prodRow As Range
    
    Set cTabble = Sheets("Saida").ListObjects(1)
    Set prodRow = insereRow(cTabble)
    
    prodRow = vet
End Sub

Sub regMovimentacao(vet As Variant)
    Dim ws As Worksheet
    Dim cTabble As ListObject
    Dim prodRow As Range
    
    Set ws = Sheets("Controle")
    Set cTabble = ws.ListObjects(1)
    Set prodRow = insereRow(cTabble)
    
    prodRow = vet
    Call remIcon_add(ws, prodRow, cTabble.ListRows.Count, 1)
End Sub

Sub removeMovim(ByVal nm As String, Optional ByVal r1 As Integer)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng1 As Range, rng2 As Range
    Dim arr() As Variant
    Dim tp As String
    Dim r2 As Integer, q As Integer
    
    Set ws = Sheets("Controle")
    Set tbl = ws.ListObjects(1)
    
    If (r1 = 0) Then r1 = Right(nm, Len(nm) - InStr(1, nm, "_"))
    
    Set rng1 = tbl.ListRows(r1).Range
    arr = rng1.Value2
    q = arr(1, 8)
    
    tp = "Saida"
    If (q > 0) Then
        tp = "Entrada"
    End If
    
    r2 = buscaMovim(tp, arr(1, 1), arr(1, 2), arr(1, 6))
    Set rng2 = Sheets(tp).ListObjects(1).ListRows(r2).Range
        
    Call ajustaIcon(ws, r1, nm)
    Call deleteRow(ws, rng1)
    Call deleteRow(Sheets(tp), rng2)

    Call atualizaEstoque(arr(1, 6), -q)

End Sub

Function buscaMovim(ws_n As String, ByVal x1 As Long, _
                  ByVal x2 As Double, ByVal x3 As Long) _
As Integer
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Integer, j As Integer
    Dim a As Integer, b As Integer
    Dim r As Integer
    Dim arr1 As Variant
    
    Set ws = Sheets(ws_n)
    Set tbl = ws.ListObjects(1)
    a = 1
    b = tbl.ListRows.Count

    arr1 = tbl.ListColumns(1).DataBodyRange.Value2

    i = bSearch_c(arr1, a, b, x1, 0)
    j = bSearch_c(arr1, i, b, x1, 1)
    
    If (i < j) Then
        arr1 = tbl.ListColumns(2).DataBodyRange.Value2
        r = bSearch_c(arr1, i, j, x2, 0)
        b = bSearch_c(arr1, r, j, x2, 1)
    Else
        buscaMovim = i
        Exit Function
    End If
    
    If (r < b) Then
        arr1 = tbl.ListColumns(6).Range.Value2
        Do While ((arr1(r + 1, 1) <> x3))
            If (r >= b) Then
                r = b
                Exit Do
            End If
            r = r + 1
        Loop
    End If
    
    buscaMovim = r
    
End Function

Function geraVetorMov(u As UserForm, nm As String, ch As String, _
                      mt As String, n_box As Integer, _
                      Optional rw As Integer = 0) As Variant
    Dim i As Integer, j As Integer, e As Integer
    Dim vet(1 To 9) As Variant
    j = 1
    e = 0

    With u
        vet(j) = CDate(.boxData)
        j = j + 1
        vet(j) = CDate(.boxHora)
        j = j + 1
        vet(j) = .boxU
        j = j + 1
        If (ch = "") Then ch = "SEM CH"
        vet(j) = ch
        If (nm = "movForm") Then
            For i = 1 To n_box
                vet(j + i) = UCase(.Controls("box" & i))
            Next
        Else
            For i = 1 To .pList.ColumnCount - 1
                vet(j + i) = .pList.List(rw, i)
            Next
        End If
        j = j + i
        vet(j) = mt
    End With
    
    geraVetorMov = vet
    
End Function

Function geraVetorMovXML(ByVal arr As Variant, ByVal nf As String, _
                         ByVal dt As Date, ByVal tm As Date) As Variant
    Dim i As Integer, j As Integer
    Dim mt As String
    Dim vet(1 To 9) As Variant
    j = 1
    mt = "PEDIDO FABRICA - NF " & nf

    vet(j) = dt
    j = j + 1
    vet(j) = tm
    j = j + 1
    vet(j) = Range("actv").Value
    j = j + 1
    vet(j) = "SEM CH"
    For i = 1 To UBound(arr)
        vet(j + i) = arr(i)
    Next
    j = j + i
    vet(j) = mt
    
    geraVetorMovXML = vet
    
End Function

Function geraVetorMvm(ByVal ch As String, ByVal pArray As Variant, _
                      ByVal qt As String, Optional ByVal tp As Integer = 0) As Variant
    Dim vet(1 To 5) As Variant
    
    vet(1) = ch
    
    vet(2) = pArray(1, 2)
    vet(3) = pArray(1, 4)
    vet(4) = pArray(1, 5)
    vet(5) = qt
    
    geraVetorMvm = vet
End Function

Function trataMotiv(u As UserForm) As String
    Dim i As Integer
    
    With u
        For i = 1 To 4
            If (u.Controls("opt_" & i) = True) Then
                trataMotiv = .Controls("opt_" & i).Caption
                Exit For
            End If
        Next
        If (i > 4) Then trataMotiv = .opt_o_txt
    End With
    
    trataMotiv = UCase(trataMotiv)
End Function

Function defineMotiv(userform_name As String, Optional ByVal tp As Integer) As String
    Dim u As Object
    Dim mU As Boolean
    Dim i As Integer
    
    If (userform_name <> "movForm") Then
        Load motivForm
        If (tp = 1) Then
            motivForm.Caption = Replace(motivForm.Caption, "do", "de Entrada no")
        Else
            motivForm.Caption = Replace(motivForm.Caption, "do", "de Saida no")
        End If
        motivForm.Show
        mU = True
        If (Not IsUserFormLoaded("motivForm")) Then Exit Function
        Set u = motivForm
        If (u.bCancel) Then
            Unload u
            Exit Function
        End If
    Else
        Set u = movForm
    End If
    
    defineMotiv = UCase(trataMotiv(u))

End Function

Function defineMotivMult(lst As Variant, ByVal tam As Integer, ByRef mt_e As String, ByRef mt_s As String) As Boolean
    Dim i As Integer, e As Integer, s As Integer
    Dim x As Variant

    defineMotivMult = False
    
    For i = 0 To tam
        If (lst(i, 4) > 0) Then
            e = e + 1
        Else
            s = s + 1
        End If
    Next
    
    If (e > 0) Then
        MsgBox "Defina uma motivacao para entrada de mercadoria", vbExclamation
        mt_e = defineMotiv("mvmForm", 1)
        If (mt_e = "") Then Exit Function
    End If
    
    If (s > 0) Then
        MsgBox "Defina uma motivacao para saida de mercadoria", vbExclamation
        mt_s = defineMotiv("mvmForm", 2)
        If (mt_s = "") Then Exit Function
    End If
    
    defineMotivMult = True

End Function

Sub removeMovimMult(ByVal pCod As Integer)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim arr As Variant
    Dim i As Integer

    Set ws = Sheets("Controle")
    Set tbl = ws.ListObjects(1)
    arr = tbl.ListColumns(6).DataBodyRange.Value2
    
    For i = UBound(arr, 1) To 1 Step -1
        If (pCod = arr(i, 1)) Then Call removeMovim("rem_" & i, i)
    Next
End Sub

Function insereDadoLista(pList As Object, v As Variant, t As Integer) As Boolean
    Dim p As Integer, i As Integer
    Dim x As Integer, c As Integer
    
    insereDadoLista = False
    With pList
        For p = 0 To .ListCount - 1
            If (c = 0) Then
                For c = 0 To .ColumnCount - 2
                    If (Len(.List(p, c)) <= 5 And IsNumeric(.List(p, c))) Then
                        Exit For
                    End If
                Next
            End If
            x = .List(p, c)
            
            If (x = v(3 - t)) Then
                Exit For
            End If
        Next
        
        If (p = .ListCount) Then
            .AddItem
            For i = 0 To .ColumnCount - 1
                .List(p, i) = v(i + 1)
            Next
        Else
            i = .ColumnCount - 1
            x = .List(p, i)
            x = x + CInt(v(5 - t))
            If (x < 0) Then
                If (Not validaEstoque(-x, getEstoque(v(3 - t)))) Then Exit Function
            ElseIf (x = 0) Then
                .RemoveItem (p)
            Else
                .List(p, i) = x
            End If
        End If
    End With
    insereDadoLista = True
End Function
