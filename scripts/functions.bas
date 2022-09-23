Attribute VB_Name = "functions"
Option Explicit

Sub scrUpdting()

    Application.ScreenUpdating = True

End Sub

Sub tglFullScreen()

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
    
    bArray = bSheet.ListObjects(1).Range.Value
    
    For c = 1 To UBound(bArray, 2)
        If (bArray(1, c) Like str) Then Exit For
    Next
    
    If (c > UBound(bArray, 2)) Then Exit Function
    
    For i = 2 To UBound(bArray, 1)
        If (bArray(i, c) = bValue) Then
            Set buscaProduto = bSheet.ListObjects(1).ListRows(i - 1).Range
            Exit For
        End If
    Next

End Function

Function bSearch_c(arr As Variant, ByVal a As Integer, ByVal b As Integer, ByVal x As Double, t As Integer) As Integer
    Dim m As Integer
    
    Select Case t
    Case 0
        While (a < b)
            m = ((a + b) / 2)
            m = m + 1 * (m > ((a + b) / 2))
            If (arr(m, 1) < x) Then
                a = m + 1
            Else
                b = m
            End If
        Wend
    Case 1
        While (a < b)
            m = ((a + b) / 2)
            m = m + 1 * (m > ((a + b) / 2))
            If (arr(m, 1) <= x) Then
                a = m + 1
            Else
                b = m
            End If
        Wend
    End Select
    
    bSearch_c = b
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
    
    aArray = tbl.HeaderRowRange.Value
    For i = 1 To tbl.HeaderRowRange.Count
        If (aArray(1, i) = "CODIGO HERDEIRO") Then
            Exit For
        End If
    Next
    
    aArray = tbl.ListColumns(i).DataBodyRange.Value
    bArray = tbl.ListColumns(tbl.ListColumns.Count).DataBodyRange.Value
    For i = 1 To tbl.ListRows.Count
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
    Case 3, 5
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

' Valida o formulario - retorna verdadeiro se valido ou falso caso contrario
' Em caso de falso, imprime em qual campo houve a invalidade
Function validaForm(uf As UserForm, nm As String, n_box As Integer) As Boolean
    Dim i As Integer
    
    validaForm = True
    For i = 1 To n_box
        If (uf.Controls("box" & i) = "") Then
            MsgBox "Informacoes de cadastro invalidas!" & _
                   vbCrLf & _
                   vbCrLf & _
                   "Insira o '" & uf.Controls("Label" & i) & "'!", vbExclamation
            uf.Controls("box" & i).SetFocus
            validaForm = False
            Exit Function
        End If
    Next
    
    If (nm = "movForm" Or nm = "mvmForm") Then
        If (uf.Controls("box" & i - 1) = 0) Then
            validaForm = False
        End If
        Exit Function
    End If
    
    If (nm = "cadForm") Then
        If (uf.box4 <= 0) Then
            MsgBox "Informacoes de cadastro invalidas!" & _
                   vbCrLf & _
                   vbCrLf & _
                   "Limite de estoque deve ser no minimo 1!", vbExclamation
            validaForm = False
        End If
        Exit Function
    End If
    
    If (Not IsError(uf.Controls("motiv_frame"))) Then
        For i = 1 To 4
            If (uf.Controls("opt_" & i)) Then
                Exit Function
            End If
        Next
        If (i > 4) Then
            If (uf.opt_o = False) Then
                MsgBox "Informacoes de cadastro invalidas!" & vbCrLf & _
                    "Defina um motivo para a movimentacao do estoque!", vbExclamation
                validaForm = False
            ElseIf (uf.opt_o_txt = "") Then
                MsgBox "Informacoes de cadastro invalidas!" & vbCrLf & _
                    "Para outras motivacoes escreva manualmente", vbExclamation
                uf.opt_o_txt.SetFocus
                validaForm = False
            End If
        End If
    End If
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
                       vbCrLf & _
                       vbCrLf & _
                       "Estoque atual do produto é " & estq, _
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
