Attribute VB_Name = "logScripts"
Option Explicit

Sub loggin()
    Dim usrRng As Range
    Set usrRng = Range("actv")
    
    logForm.Show
    While (usrRng.Value = "")
        MsgBox "Necessario logar para acessar a planilha!", vbExclamation
        logForm.Show
    Wend

End Sub

Sub loggin_A()
    Dim ws As Worksheet
    Dim usrRng As Range
    Dim user As Object, pass As Object, key As String
    Set ws = Sheets("Acesso")
    Set usrRng = Range("actv")
    Set user = ws.OLEObjects("TextBox1").Object
    Set pass = ws.OLEObjects("TextBox2").Object
    
    key = buscaAcesso(user.Value)
    
    If (key <> pass.Value) Then
        MsgBox "Usuario/senha invalidos!"
    Else
        Call planAccess(user.Value)
        user.Value = ""
        pass.Value = ""
        ws.Shapes("logginStyle").Line.Visible = msoFalse
    End If
    
End Sub

Sub logout()
    Dim ws As Worksheet

    Application.ScreenUpdating = False
    If (TypeName(Application.Caller) = "String") Then
        If (MsgBox("Encerrar sua sessão?", vbQuestion + vbYesNo) = vbNo) Then Exit Sub
    End If
    For Each ws In ActiveWorkbook.Worksheets
        If (ws.Name = "Acesso") Then
            ws.Visible = xlSheetVisible
            ws.Activate
            ws.OLEObjects("TextBox1").Object = ""
            ws.OLEObjects("TextBox2").Object = ""
            ws.Shapes("logginStyle").Line.Visible = msoFalse
        Else
            ws.Visible = xlSheetVeryHidden
        End If
    Next
    Range("actv") = ""
    Application.ScreenUpdating = True
    
End Sub

Function buscaAcesso(user As String) As String
    Dim i As Integer
    Dim tbl As ListObject
    Set tbl = Sheets("Usuarios").ListObjects(1)
    
    buscaAcesso = "-1"
    For i = 1 To tbl.ListRows.Count + 1
        If (tbl.Range(i, 1) = user) Then
            buscaAcesso = tbl.Range(i, 2)
            Exit For
        End If
    Next
End Function

Sub anotherPass()
    With logForm
        If (.aviso.Visible = True) Then
            .aviso.Visible = False
            .Height = .Height - 20
            .aviso.Top = 140
            .logginBtn.Top = .logginBtn.Top - 20
            .cancelBtn.Top = .cancelBtn.Top - 20
        End If
    End With
End Sub

Sub invalidPass()
    With logForm
        .Height = .Height + 20
        .aviso.Top = .logginBtn.Top - 4
        .logginBtn.Top = .logginBtn.Top + 20
        .cancelBtn.Top = .cancelBtn.Top + 20
        .aviso.Visible = True
    End With
End Sub

Sub planAccess(user As String)
    Dim ws As Worksheet
    
    Unload logForm
    
    Application.ScreenUpdating = False
    For Each ws In ActiveWorkbook.Sheets
        If (ws.Name <> "Usuarios" And ws.Name <> "empty") Then
            ws.Visible = xlSheetVisible
        End If
    Next
    If (user = "admin") Then
        Sheets("Usuarios").Visible = xlSheetVisible
        Sheets("Acesso").Visible = xlSheetHidden
        ActiveWindow.DisplayWorkbookTabs = True
    Else
        Sheets("Acesso").Visible = xlSheetVeryHidden
        ActiveWindow.DisplayVerticalScrollBar = True
    End If
    Range("actv") = UCase(user)
    Application.ScreenUpdating = True
    Application.CalculateFull
    
    MsgBox "Bem vindo " & UCase(user) & "!"
    
End Sub
