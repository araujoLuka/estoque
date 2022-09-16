Attribute VB_Name = "functions"
Option Explicit

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

Function buscaProduto(ByVal bIndex As Integer, ByVal bValue As Variant, Optional bSheet As Worksheet) As Range
    Dim bArray() As Variant
    Dim bTam As Integer, i As Integer
    
    If (bValue = "") Then Exit Function
    
    If (IsNumeric(bValue)) Then
        If (Len(bValue) >= 5) Then
            bValue = CCur(bValue)
        Else
            bValue = CInt(bValue)
        End If
        bIndex = 2 * bIndex
    Else
        bValue = UCase(bValue)
    End If
            
    If (bSheet Is Nothing) Then
        Set bSheet = Sheets("Cadastro")
    Else
        bIndex = WorksheetFunction.Ceiling_Math((bIndex + 1) / 2)
    End If
    bArray = bSheet.ListObjects(1).ListColumns(bIndex).DataBodyRange.Value
    bTam = bSheet.ListObjects(1).Range.Rows.Count
    
    For i = 1 To bTam - 1
        If (bArray(i, 1) = bValue) Then
            Set buscaProduto = bSheet.ListObjects(1).ListRows(i).Range
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

Function validaMovim(cHerdeiro As String, tipo As Integer) As Boolean
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

Sub travaCampo(c As Control, Optional x As String = "")
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
