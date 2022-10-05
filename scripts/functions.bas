Attribute VB_Name = "functions"
Option Explicit

Sub scrUpdting()

    Application.ScreenUpdating = True

End Sub

Sub tglFullScreen()
Attribute tglFullScreen.VB_ProcData.VB_Invoke_Func = "T\n14"

    Application.DisplayFullScreen = Not Application.DisplayFullScreen

End Sub

Function trataCaller(cllr As Variant, ByRef nm As String) As Integer
    Dim rw As Integer
    
    rw = 0
    
    If (Not IsError(cllr)) Then
        nm = cllr
        rw = Right(nm, Len(nm) - InStr(1, nm, "_"))
    End If
    
    trataCaller = rw

End Function

Function identificaTipo(criterio As Integer) As Integer
    Dim i As Integer
    
    If (IsNumeric(criterio)) Then
        If (IsError(CInt(criterio))) Then
            i = 1
        Else
            i = 2
        End If
    Else
        i = 3
    End If
    
    identificaTipo = i

End Function

Function buscaProduto(ByVal bValue As Variant, ByVal bType As Integer, Optional bSheet As Worksheet) As Range
    Dim bArray() As Variant
    Dim i As Integer, c As Integer
    Dim str As String
    
    If (bValue = "") Then Exit Function
    
    Select Case bType
    Case 1
        str = "*BARRAS"
    Case 2
        str = "*INTERNO"
    Case 3
        str = "*BARRAS"
    End Select
    
    If (IsNumeric(bValue)) Then
        If (Len(bValue) >= 5) Then
            bValue = CCur(bValue)
        Else
            bValue = CInt(bValue)
        End If
    Else
        bValue = UCase(bValue)
    End If
    
    If (bSheet Is Nothing) Then Set bSheet = Sheets("Cadastro")
    
    c = defineColuna(bSheet.ListObjects(1), str)
    If (c = 0) Then Exit Function
    
    bArray = bSheet.ListObjects(1).ListColumns(c).Range.Value2
        
    For i = 2 To UBound(bArray, 1)
        If (bArray(i, 1) = bValue) Then
            Set buscaProduto = bSheet.ListObjects(1).ListRows(i - 1).Range
            Exit For
        End If
    Next

End Function

Function bSearch_c(arr As Variant, ByVal a As Integer, ByVal b As Integer, ByVal x As Double, t As Integer) As Integer
    Dim m As Integer, mx As Integer
    mx = b
    
    Select Case t
    Case 0 'Procura a primeira ocorrencia
        Do While (a < b)
            m = ((a + b) / 2)
            m = m + 1 * (m > ((a + b) / 2))
            If (arr(m, 1) < x) Then
                a = m + 1
            Else
                b = m
            End If
        Loop
    Case 1 'Procura a ultima ocorrencia
        Do While (a < b)
            m = ((a + b) / 2)
            m = m + 1 * (m > ((a + b) / 2))
            If (arr(m, 1) <= x) Then
                If (m = mx) Then Exit Do
                If (arr(m + 1, 1) > x) Then Exit Do
                a = m + 1
            Else
                b = m
            End If
        Loop
    End Select
    
    If (a < b) Then
        bSearch_c = m
    Else
        bSearch_c = b
    End If
End Function

Function countFormTBX(mForm As UserForm) As Integer
    Dim cCont As Control
    Dim i As Integer
    i = 0

    For Each cCont In mForm.Controls
        If (TypeName(cCont) = "TextBox") And cCont.Name = "box" & i + 1 Then
            i = i + 1
        End If
    Next cCont
    
    countFormTBX = i
End Function

Function trataCodigo(entry As String, ByRef index As Integer, Optional ByRef herd As Boolean) As String
    Dim cod As String
    Dim x As Long
    
    Select Case Len(entry)
    Case 1 To 5
        cod = entry
        index = 2
    Case 12
        cod = Left(entry, 3)
        herd = True
        index = 2
    Case 13
        cod = entry
        index = 1
    Case 14
        cod = Left(entry, 1) - 6 & Right(Left(entry, 5), 4)
        herd = True
        index = 2
    Case 16
        x = Bin2Dec(Left(entry, 3))
        cod = Right(entry, 13) + x
        If (IsUserFormLoaded("movForm")) Then
            movForm.box4 = x + 1
        End If
        index = 1
    Case Else
        cod = ""
    End Select
    trataCodigo = cod

End Function

Function validaMovim(ByVal cHerdeiro As String, tipo As Integer) As Boolean
    Dim tbl As ListObject
    Dim aArray() As Variant, bArray() As Variant
    Dim i As Integer, m As Integer, x As Integer
    Dim acao As String
    Set tbl = Sheets("Controle").ListObjects(1)
    validaMovim = True
    x = 0
    
    If (Not IsNumeric(cHerdeiro)) Then Exit Function
    
    aArray = tbl.HeaderRowRange.Value2
    For i = 1 To tbl.HeaderRowRange.Count
        If (aArray(1, i) = "CODIGO HERDEIRO") Then
            Exit For
        End If
    Next
    
    aArray = tbl.ListColumns(i).Range.Value2
    bArray = tbl.ListColumns(tbl.ListColumns.Count).Range.Value
    For i = 2 To tbl.ListRows.Count
        If (aArray(i, 1) = CCur(cHerdeiro)) Then
            x = x + bArray(i, 1)
        End If
    Next
    
    If (x <> 0) Then
        If (tipo = 1) Then
            m = 1
            acao = "adicionado ao"
        ElseIf (tipo = 2) Then
            m = -1
            acao = "subtraido do"
        End If
        If (x * m > 0) Then
            MsgBox "Produto de codigo herdeiro " & cHerdeiro & _
                    " ja " & acao & " estoque!", vbExclamation
            validaMovim = False
        End If
    End If

End Function

' Valida o formulario - retorna verdadeiro se valido ou falso caso contrario
' Em caso de falso, imprime em qual campo houve a invalidade
Function validaForm(uf As UserForm, nm As String, n_box As Integer) As Boolean
    Dim i As Integer
    
    validaForm = False
    For i = 1 To n_box
        If (uf.Controls("box" & i) = "") Then
            MsgBox "Informacoes de cadastro invalidas!" & _
                   vbCrLf & _
                   vbCrLf & _
                   "Insira o '" & uf.Controls("Label" & i) & "'!", vbExclamation
            uf.Controls("box" & i).SetFocus
            Exit Function
        End If
        If (uf.Controls("Label" & i) <> "Produto") Then
            If (Not IsNumeric(uf.Controls("box" & i)) And Not uf.Controls("box" & i) Like "SEM*") Then
                MsgBox "Informacoes de cadastro invalidas!" & _
                       vbCrLf & _
                       vbCrLf & _
                       "O campo '" & uf.Controls("Label" & i) & "' deve ser numerico!", vbExclamation
                Exit Function
            End If
        Else
            If (IsNumeric(uf.Controls("box" & i))) Then
                MsgBox "Informacoes de cadastro invalidas!" & _
                       vbCrLf & _
                       vbCrLf & _
                       "O campo '" & uf.Controls("Label" & i) & "' não pode ser numérico!", vbExclamation
                Exit Function
            ElseIf (IsNumeric(Left(uf.Controls("box" & i), 1))) Then
                MsgBox "Informacoes de cadastro invalidas!" & _
                       vbCrLf & _
                       vbCrLf & _
                       "O campo '" & uf.Controls("Label" & i) & "' não pode iniciar com um número!", vbExclamation
                Exit Function
            End If
        End If
    Next
    
    If (nm = "movForm" Or nm = "mvmForm") Then
        If (uf.Controls("box" & i - 1) = 0) Then
            MsgBox "Impossivel adicionar/subtrair 0(zero) unidades!", vbExclamation
            Exit Function
        End If
    ElseIf (nm = "cadForm") Then
        If (uf.box4 <= 0) Then
            MsgBox "Informacoes de cadastro invalidas!" & _
                   vbCrLf & _
                   vbCrLf & _
                   "Limite de estoque deve ser no minimo 1!", vbExclamation
            Exit Function
        End If
    End If
    validaForm = True

End Function

Function validaMotiv(str As String, Optional errors As Integer) As Boolean
    Dim ret As Boolean
    
    ret = True
    If (str = "") Then
        MsgBox "Motivo invalido para movimentacao do estoque!"
        If (errors < 3) Then
            errors = errors + 1
        Else
            MsgBox "Limite maximo de erros atingido!" & vbCrLf & _
                   "Movimentacao abortada", vbCritical
        End If
        ret = False
    End If
    
    validaMotiv = ret

End Function

Function validaEstoque(ByVal qtd As Integer, ByVal estq As Integer)
    Dim ret As Boolean
    
    ret = True
    If (qtd > estq) Then
        MsgBox Prompt:="Quantidade deve ser " & _
                       "menor ou igual ao total disponivel em estoque" & _
                       vbCrLf & vbCrLf & _
                       "Estoque atual do produto: " & estq, _
               Buttons:=vbExclamation, _
               Title:="Falha ao registrar saida de mercadoria"
        ret = False
        Exit Function
    End If
    
    validaEstoque = ret
End Function

Function IsUserFormLoaded(ByVal UFName As String) As Boolean
    Dim UForm As Object
    
    IsUserFormLoaded = False
    For Each UForm In VBA.UserForms
        If UForm.Name = UFName Then
            IsUserFormLoaded = True
            Exit For
        End If
    Next
End Function

Sub travaCampo(c As Control, Optional ByVal x As String = "")
    With c
        .Value = x
        .BackColor = &H80000016
        .Enabled = False
    End With
End Sub

Sub destravaCampo(c As Control)
    With c
        .BackColor = vbWhite
        .ForeColor = &H80000008
        .Enabled = True
        .SetFocus
    End With
End Sub

Function Bin2Dec(sMyBin As String) As Long
    Dim x As Integer
    Dim iLen As Integer

    iLen = Len(sMyBin) - 1
    For x = 0 To iLen
        Bin2Dec = Bin2Dec + _
          Mid(sMyBin, iLen - x + 1, 1) * 2 ^ x
    Next
End Function

Sub highlightSelection(ByVal Target As Range)
    Dim rw As Integer
    Dim fillColor As MsoColorType, bordColor As MsoColorType
    fillColor = RGB(230, 230, 230)
    bordColor = RGB(0, 176, 80)

    rw = Target.Row - Target.ListObject.HeaderRowRange.Row
    
    If (Target.ListObject.ListRows.Count = 0 Or rw <= 0) Then Exit Sub

    With Target.ListObject.ListRows(rw)
        With .Range
            .Interior.Color = fillColor
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Color = bordColor
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Color = bordColor
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Color = bordColor
                .Weight = xlMedium
            End With
        End With
    End With
End Sub

Function defineColuna(ByVal tbl As ListObject, ByVal nm As String) As Integer
    Dim i As Integer
    Dim arr As Variant
    
    arr = tbl.HeaderRowRange.Value2
    
    For i = 1 To UBound(arr, 2)
        If (arr(1, i) Like nm) Then Exit For
    Next
    
    defineColuna = i
    If (i > UBound(arr, 2)) Then defineColuna = 0
End Function

Sub limparPlanilha()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Integer, tam As Integer
    
    Set ws = Sheets("Cadastro")
    Set tbl = ws.ListObjects(1)
    
    tam = tbl.ListRows.Count
    
    If (tam = 0) Then Exit Sub
    
    For i = tam To 1 Step -1
        Call removeProduto(tbl.ListRows(i).Range, True, True)
    Next
End Sub

