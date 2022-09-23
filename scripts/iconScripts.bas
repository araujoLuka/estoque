Attribute VB_Name = "iconScripts"
' Scripts - Icones de Controle
'
' Autor: Lucas Araujo
'
' Resumo: Modulo com procedimentos para adicionar e remover icones de
'         controle das planilhas e para realizar tarefas como editar
'         um produto ou remover algum registro
'--------------------------------------------------------------------'

' \\\ Forca declaracao de variaveis no modulo atual ///
Option Explicit

' \\\ Constantes que sao usadas no modulo ///
' Caminhos de imagens para icones
Private Const EDIT_PATH As String = "\resources\edit_icon1.png"
Private Const REM_PATH As String = "\resources\rem_icon2.png"

' Espaco entre icones
Private Const IC_SPACE As Integer = 15
'--------------------------------------------------------------------'

' \\\ Funcoes gerais ///

' Adiciona os icones de controle em cada linha de tabela da planilha atual
Sub addIcons()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim rw As Integer
    
    Set ws = ActiveSheet    ' Variavel com a planilha atual
    Set tbl = ws.ListObjects(1)     ' Variavel com a tabela da planilha
    
    ' Desativa a atualizacao de tela para melhorar desempenho da execucao
    Application.ScreenUpdating = False
    
    ' Loop para tratar cada linha da tabela
    For rw = 1 To tbl.ListRows.Count
        Set rng = tbl.ListRows(rw).Range
        
        ' Seletor para definir o tipo de inclusao de icone
        Select Case ws.Name
            Case "Cadastro"
                Call editIcon_add(ws, rng, rw, 1)
                Call remIcon_add(ws, rng, rw, 2)
        
            Case "Controle"
                Call remIcon_add(ws, rng, rw, 1)
            
            Case Else
                MsgBox "Ainda nao ha utilidade de icones para a planilha atual"
                Exit For
                
        End Select
    Next
    
    ' Reativa a atualizacao de tela
    Application.ScreenUpdating = True
    
End Sub

' Deleta todos os icones de controle da planilha atual
Sub deleteIcons()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rw As Integer
    
    Set ws = ActiveSheet    ' Variavel com a planilha atual
    Set tbl = ws.ListObjects(1)     ' Variavel com a tabela da planilha
    
    ' Loop para tratar cada linha
    For rw = 1 To tbl.ListRows.Count
        If (deleteEditIcon(ws, rw) = False And deleteRemIcon(ws, rw) = False) Then Exit For
    Next
    
End Sub

' Atualiza os icones, deletando todos e adicionando novamente
' *** apenas para planilha atual
Sub updateIcons()
    Call deleteIcons
    Call addIcons
End Sub
'--------------------------------------------------------------------'

' Ajusta o nome dos icones da linha 'i' para baixo. Para caso de exclus�o de linha
' Recebe a planilha onde havera exclusao, o numero da linha e o nome do icone da linha
' ***Em caso de duvida, existe um exemplo ao fim do procedimento
Sub ajustaIcon(ws As Worksheet, ByVal i As Integer, ByVal nm As String)
    Dim aux As Shape, sh As Shape
    Dim iType As String
    
    ' Salva o icone da linha que sera excluido
    Set sh = ws.Shapes(nm)
    
    ' Define o tipo de icone (edicao='edit_' ou remocao='rem_')
    iType = Left(nm, InStr(1, nm, "_"))
    
    For i = i To ws.ListObjects(1).ListRows.Count - 1
        nm = iType & i + 1
        
        On Error Resume Next
        Set aux = ws.Shapes(nm)
        On Error GoTo 0
        
        If (aux Is Nothing) Then
            Exit For
        Else
            aux.Name = sh.Name
            sh.Name = nm
            Set aux = Nothing
        End If
    Next
    
    If (iType = "rem_") Then
        Call deleteRemIcon(ws, i)
    Else
        Call deleteEditIcon(ws, i)
    End If

'------------------------------------------------------------------------------
'[1]
'> Exemplo {
'       Uma tabela tem 4 linhas e deseja excluir a linha 2.
'       Icone da linha deve ser excluido
'
'       Entao, icone debaixo recebe o nome do icone atual e assim sucessivamente
'
'       No final, Icones 3 e 4 seram 2 e 3 e o Icone da linha eh excluido.
'}

End Sub

' \\\ Funcoes - Icone de Edicao ///

' Adiciona um icone de edicao
Sub editIcon_add(ws As Worksheet, ByVal rng As Range, ByVal rw As Integer, pos As Integer)
    Dim x As Single, y As Single, size As Single
    Dim pic As Shape
    Dim imgPath As String

    ' Define a celula aonde o icone ira ser inserido
    Set rng = rng(1, rng.Columns.Count).Offset(0, 1)
    
    ' Define o tamanho do icone, baseado na celula definida anteriormente
    size = rng.Height * 0.65

    ' Define a posicao (x e y) do icone
    x = rng.Left + pos * IC_SPACE + pos * size - size
    y = rng.Top + rng.Height / 2 - size / 2
    
    ' Define o caminho da imagem do icone
    imgPath = ActiveWorkbook.Path & EDIT_PATH
    
    ' Inclui o icone com os dados anteriores
    Set pic = ws.Shapes.AddPicture(imgPath, False, True, x, y, size, size)
    pic.Name = "edit" & "_" & rw    ' Nomeia o icone ("edit_" + 'n� da linha')
    
    ' Define a acao de clique do icone
    pic.OnAction = "'" & ActiveWorkbook.Name & "'!" & "iniciaAtualiz"
    
End Sub

Function deleteEditIcon(ws As Worksheet, ByVal rw As Integer) As Boolean
    Dim i As Integer
    Dim nm As String
    Dim toDel As Shape
    
    i = rw
    nm = "edit_" & i
    
    deleteEditIcon = False
    
    On Error Resume Next
    Set toDel = ws.Shapes(nm)
    On Error GoTo 0
    
    If (toDel Is Nothing) Then Exit Function
    
    toDel.Delete
    deleteEditIcon = True
    
End Function

' Adiciona um icone de remocao
Sub remIcon_add(ws As Worksheet, ByVal rng As Range, ByVal rw As Integer, pos As Integer)
    Dim x As Single, y As Single, size As Single
    Dim pic As Shape
    Dim imgPath As String
    
    Set rng = rng(1, rng.Columns.Count).Offset(0, 1)
    size = rng.Height * 0.65

    x = rng.Left + pos * (IC_SPACE + size) - size
    y = rng.Top + (rng.Height / 2) - (size / 2.3)
    
    imgPath = ActiveWorkbook.Path & REM_PATH
    
    Set pic = ws.Shapes.AddPicture(imgPath, False, True, x, y, size, size)
    pic.Name = "rem" & "_" & rw
    
    pic.OnAction = "'" & ActiveWorkbook.Name & "'!" & "remIcon_event"
    
End Sub

Function deleteRemIcon(ws As Worksheet, ByVal rw As Integer) As Boolean
    Dim i As Integer
    Dim aux As String
    Dim toDel As Shape
    
    i = rw
    aux = "rem_" & i
    
    deleteRemIcon = False
    
    On Error Resume Next
    Set toDel = ws.Shapes(aux)
    On Error GoTo 0
    
    If (toDel Is Nothing) Then Exit Function
    
    toDel.Delete
    deleteRemIcon = True
    
End Function

Sub remIcon_event()
    Dim ws As Worksheet
    Dim nm As String
    Dim rw As Integer
    Dim mbResult As VbMsgBoxResult
    
    Set ws = ActiveSheet
    
    rw = trataCaller(Application.Caller, nm)
    If (rw = 0) Then Exit Sub

    Select Case ws.Name
        Case "Cadastro"
            Union(Range("A1"), ws.ListObjects(1).ListRows(rw).Range).Select
            mbResult = MsgBox("Deseja excluir registro selecionado?", vbYesNo, "Exclusao de cadastro?")
            If (mbResult = vbYes) Then Call removeProduto(ws.ListObjects(1).ListRows(rw).Range)
            
        Case "Controle"
            Union(Range("A1"), ws.ListObjects(1).ListRows(rw).Range).Select
            mbResult = MsgBox("Deseja excluir registro selecionado?", vbYesNo, "Exclusao de movimenta��o?")
            If (mbResult = vbYes) Then Call remMov(nm, rw)
            
        Case Else
            MsgBox "Ainda nao tem funcao para essa planilha..."
            
    End Select
    
    Range("A1").Select
End Sub
