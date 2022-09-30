VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Scripts - Planilha Acesso
'
' Autor: Lucas Araujo
'
' Resumo: Planilha que controla login na planilha de Estoque
'---------------------------------------------------------------'

' \\\ Forca declaracao de variaveis ///
Option Explicit

' \\\ Controle da planilha ///

'Comandos ao acessar a planilha
Private Sub Worksheet_Activate()
        
    ' Desprotege a planilha para fazer alteracao
    Unprotect
    
    ' Seleciona o intervalo para zoom dinamico
    Range("view").Select
    With ActiveWindow
        .Zoom = True    ' Zoom dinamico no intervalo selecionado
        .DisplayHorizontalScrollBar = False 'Remove a barra de rolagem horizontal
        .DisplayVerticalScrollBar = False   'Remove a barra de rolagem vertical
        .DisplayWorkbookTabs = False        'Remove a aba de planilhas
    End With
    TextBox1.Activate
    
    ' Protege a planilha novamente
    Protect

End Sub

'Comandos ao efetuar loggin
Private Sub Worksheet_Deactivate()
    ' Limpa os campos de acesso
    TextBox1.Value = ""
    TextBox2.Value = ""
    
    ' Reseta o estilo do botao de acesso para o padrao
    Shapes("logginStyle").Line.Visible = msoFalse
    
End Sub
'-------------------------------------------'

' \\\ Controle de selecao dos campos de login /// '

'Controle do campo de usuario
Private Sub TextBox1_GotFocus()
    ' Seleciona todo o texto ao selecionar o campo
    With TextBox1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'Controle do campo de senha
Private Sub TextBox2_GotFocus()
    ' Seleciona todo o texto ao selecionar o campo
    With TextBox2
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
'-------------------------------------------'

' \\\ Controle de atalhos nos campos/botao /// '

'Controle do campo de usuario
Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Verifica o TAB
    ' TAB -> Campo de senha
    ' SHIFT + TAB -> Botao de login
    If (Shift) Then
        If (KeyCode = vbKeyTab) Then logginBtn.Activate
    Else
        If (KeyCode = vbKeyTab) Then TextBox2.Activate
    End If
    
    ' Verifica o ENTER
    If (KeyCode = vbKeyReturn) Then Call loggin_A
End Sub

'Controle do campo de senha
Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Verifica o TAB
    ' TAB -> Botao de login
    ' SHIFT + TAB -> Campo de usuario
    If (Shift) Then
        If (KeyCode = vbKeyTab) Then TextBox1.Activate
    Else
        If (KeyCode = vbKeyTab) Then logginBtn.Activate
    End If
    
    ' Verifica o ENTER
    If (KeyCode = vbKeyReturn) Then Call loggin_A
End Sub

'Controle do botao de login
Private Sub logginBtn_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Verifica o TAB
    ' TAB -> Campo de usuario
    ' SHIFT + TAB -> Campo de senha
    If (Shift) Then
        If (KeyCode = vbKeyTab) Then TextBox2.Activate
    Else
        If (KeyCode = vbKeyTab) Then TextBox1.Activate
    End If
    
    ' Verifica o ENTER
    If (KeyCode = vbKeyReturn) Then Call loggin_A
End Sub
'-------------------------------------------'

' \\\ Animacao do botao de loggin ///'

' Faz um contorno na forma quando selecionado
Private Sub logginBtn_GotFocus()
    With Shapes("logginStyle").Line
        .Visible = msoTrue
    End With
End Sub

' Remove o contorno ao tirar a selecao
Private Sub logginBtn_LostFocus()
    With Shapes("logginStyle").Line
        .Visible = msoFalse
    End With
End Sub
'-------------------------------------------'
