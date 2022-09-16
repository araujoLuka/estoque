Attribute VB_Name = "movScripts"
' Autor: Lucas Araujo
' Ultima atualizacao: 09/07/2022
'
' Modulo para procedimentos do formulario de movimenta��o de estoque, 'movForm'

Option Explicit

' Inicializa a movimentacao de estoque
Sub iniciaMovimentacao()
    
    On Error Resume Next
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

Sub atualizaEstoque(pRng As Range, qtd As Variant)
    Dim ws As Worksheet
    Dim eRng As Range
    Dim i As Integer
    
    Set ws = Sheets("Estoque")
    Set eRng = buscaProduto(2, pRng(1, 4), ws)
    
    qtd = CInt(qtd)
    
    For i = 1 To eRng.Count
        If (eRng.ListObject.HeaderRowRange(1, i) = "ESTOQUE") Then
            Exit For
        End If
    Next
    
    eRng.Cells(1, i) = eRng.Cells(1, i) + qtd
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
    Call addRemIcon(ws, prodRow, cTabble.ListRows.Count, 1)
End Sub

Sub remMov()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng1 As Range, rng2 As Range
    Dim arr() As Variant
    Dim nm As String, tp As String
    Dim r1 As Integer, r2 As Integer
    Dim i As Integer, j As Integer
    Dim q As Integer
    
    Set ws = ActiveSheet
    Set tbl = ws.ListObjects(1)
    
    If (IsError(Application.Caller)) Then Exit Sub
    nm = Application.Caller
    r1 = Right(nm, Len(nm) - InStr(1, nm, "_"))
    Set rng1 = tbl.ListRows(r1).Range
    arr = rng1.Value2
    q = arr(1, 8)
    
    If (q > 0) Then
         tp = "Entrada"
    Else
         tp = "Saida"
    End If
    
    r2 = buscaMov(tp, arr(1, 1), arr(1, 2), arr(1, 6))
    Set rng2 = Sheets(tp).ListObjects(1).ListRows(r2).Range
    
    Application.ScreenUpdating = False
    
    Call ajustaIcon(ws, r1, nm)
    Call deleteRemIcon(ws, r1)
    Call deleteRow(ws, rng1, 1, nm)
    Call deleteRow(Sheets(tp), rng2, 0)

    Call atualizaEstoque(buscaProduto(2, arr(1, 6)), -q)

    Application.ScreenUpdating = True

End Sub

Function buscaMov(ws_n As String, ByVal x1 As Long, _
                  ByVal x2 As Double, ByVal x3 As Long) _
As Integer
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Integer, j As Integer
    Dim a As Integer, b As Integer
    Dim r As Integer
    Dim arr1() As Variant
    
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
        r = i
    End If
    
    arr1 = tbl.ListColumns(6).DataBodyRange.Value2
    While ((arr1(r, 1) <> x3) And r <= b)
        r = r + 1
    Wend
    
    buscaMov = r
    
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

Function geraVetorMvm(ch As String, pArray As Variant, qt As String) As Variant
    Dim vet(0, 0 To 4) As Variant
    
    vet(0, 0) = ch
    vet(0, 1) = pArray(1, 2)
    vet(0, 2) = pArray(1, 4)
    vet(0, 3) = pArray(1, 5)
    vet(0, 4) = qt
    
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
End Function

Function defineMotiv(u As Object) As String
    Dim mU As Boolean
    Dim i As Integer, e As Integer
    Dim str As String
    e = 0
    
    If (u.Name <> "movForm") Then
        motivForm.Show
        mU = True
        Set u = motivForm
        If (u.bCancel) Then
            Unload u
            Exit Function
        End If
    End If
    
    Do
        str = trataMotiv(u)
        If (str = "") Then
            MsgBox "Motivo invalido para movimenta��o do estoque!"
            e = e + 1
            If (mU) Then u.Show
        Else
            str = UCase(str)
            defineMotiv = str
            If (mU) Then Unload u
        End If
    Loop Until str <> "" Or e >= 2

    If (e >= 2) Then MsgBox "Limite maximo de erros atingido!" & _
                        vbCrLf & "Movimenta��o abortada", vbCritical

End Function

Sub insereDadoLista(pList As Object, v As Variant)
    Dim p As Integer, i As Integer
    Dim x As Long
    
    With pList
        For p = 0 To .ListCount - 1
            x = .List(p, 2)
            If (x = v(0, 2)) Then
                Exit For
            End If
        Next
        
        If (p = .ListCount) Then
            .AddItem
            For i = 0 To .ColumnCount - 1
                .List(p, i) = v(0, i)
            Next
        Else
            i = .ColumnCount - 1
            x = .List(p, i)
            .List(p, i) = x + CInt(v(0, i))
            x = .List(p, i)
            If (x = 0) Then
                .RemoveItem (p)
            End If
        End If
    End With
End Sub
