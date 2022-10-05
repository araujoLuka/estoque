Attribute VB_Name = "logScripts"
Option Explicit

Sub iniciaLoggin()
    Dim usrRng As Range
    Set usrRng = Range("actv")
    
    logginForm.Show
    While (usrRng.Value = "")
        MsgBox "Necessario logar para acessar a planilha!", vbExclamation
        logginForm.Show
    Wend

End Sub

Sub loggin_A(ByVal user As String, ByVal pass As String)
    Dim ws As Worksheet
    Dim rng As Range
    Dim arr As Variant
    
    Set ws = Sheets("Acesso")
    
    If (user = "") Then
        MsgBox "Usuario/senha invalidos!"
        Exit Sub
    End If
        
    Set rng = buscaAcesso(user)
    
    If (rng Is Nothing) Then
        MsgBox "Usuario/senha invalidos!"
        Exit Sub
    End If
        
    arr = rng.Value
    
    If (arr(1, 2) <> pass) Then
        MsgBox "Usuario/senha invalidos!"
        Exit Sub
    End If
    
    Call planAccess(arr, rng(1, 4))
    
End Sub

Sub logout()
    Dim ws As Worksheet

    ThisWorkbook.Activate
    Application.ScreenUpdating = False
    If (TypeName(Application.Caller) = "String") Then
        If (MsgBox("Encerrar sua sessão?", vbQuestion + vbYesNo) = vbNo) Then Exit Sub
    End If
    For Each ws In ThisWorkbook.Worksheets
        If (ws.Name = "Acesso") Then
            ws.Visible = xlSheetVisible
            ws.Activate
            ws.OLEObjects("TextBox1").Object.Value = ""
            ws.OLEObjects("TextBox2").Object.Value = ""
            ws.Shapes("logginStyle").Line.Visible = msoFalse
        Else
            ws.Visible = xlSheetVeryHidden
        End If
    Next
    Range("actv") = ""
    Application.ScreenUpdating = True
    
End Sub

Function buscaAcesso(user As String) As Range
    Dim i As Integer
    Dim tbl As ListObject
    Dim arr As Variant
    
    Set tbl = Sheets("Usuarios").ListObjects(1)
    arr = tbl.DataBodyRange.Value2
    
    Set buscaAcesso = Nothing
    For i = 1 To UBound(arr, 1)
        If (arr(i, 1) = LCase(user)) Then
            Set buscaAcesso = tbl.ListRows(i).Range
            Exit For
        End If
    Next
End Function

Sub anotherPass()
    With logginForm
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
    With logginForm
        .Height = .Height + 20
        .aviso.Top = .logginBtn.Top - 4
        .logginBtn.Top = .logginBtn.Top + 20
        .cancelBtn.Top = .cancelBtn.Top + 20
        .aviso.Visible = True
    End With
End Sub

Sub planAccess(arr As Variant, lastAccs As Range)
    Dim ws As Worksheet
    Dim usrRange As Range
    Dim user As String
    
    Set usrRange = Range("actv")
    
    Unload logginForm
    
    Application.ScreenUpdating = False
    
    For Each ws In ActiveWorkbook.Sheets
        If (ws.Name <> "Usuarios" And ws.Name <> "empty") Then
            ws.Visible = xlSheetVisible
        End If
    Next
    
    If (arr(1, 3) = 3 Or arr(1, 1) = "admin") Then
        Sheets("Usuarios").Visible = xlSheetVisible
        Sheets("Acesso").Visible = xlSheetHidden
        Sheets("Acesso").Unprotect
        ActiveWindow.DisplayWorkbookTabs = True
    Else
        Sheets("Acesso").Visible = xlSheetVeryHidden
        ActiveWindow.DisplayVerticalScrollBar = True
    End If
    usrRange = UCase(arr(1, 1))
    
    If (lastAccs.Value <> Date) Then
        If (arr(1, 1) = "admin") Then
            MsgBox "Bem vindo Administrador!", vbInformation, "Mensagem de Boas-Vindas"
        Else
            user = UCase(Left(arr(1, 1), 1)) & Right(arr(1, 1), Len(arr(1, 1)) - 1)
            MsgBox "Bem vindo " & user & "!", vbInformation, "Mensagem de Boas-Vindas"
        End If
    End If
    lastAccs = Date
    
    Application.ScreenUpdating = True
    Application.CalculateFull
    
End Sub
