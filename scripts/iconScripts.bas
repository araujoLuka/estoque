Attribute VB_Name = "iconScripts"
Option Explicit

Private Const EDIT_PATH As String = ".\resources\edit_icon1.png"
Private Const REM_PATH As String = ".\resources\rem_icon2.png"
Private Const IC_SPACE As Integer = 15

Sub multi_addRemIcon()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim sh As Shape
    Dim c As Integer, r As Integer
    Set ws = ActiveSheet
    Set tbl = ws.ListObjects(1)
    c = tbl.ListColumns.Count
    r = 1

    For Each sh In ws.Shapes
        If (sh.Name = "rem_" & r) Then r = r + 1
    Next
    
    Set sh = ws.Shapes("rem_0")
    
    For r = r To tbl.ListRows.Count
        Set rng = tbl.ListRows(r).Range(1, c)
        With sh.Duplicate
            .Name = "rem_" & r
            .Left = rng.Left + (rng.offset(0, 1).Left - rng.Left) / 2 - sh.Width / 3
            .Top = rng.Top + (rng.offset(1, 0).Top - rng.Top) / 2 - sh.Height / 2
        End With
    Next
    
End Sub

Sub addIcons()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim rw As Integer
    
    Set ws = ActiveSheet
    Set tbl = ws.ListObjects(1)
    
    Application.ScreenUpdating = False
    
    For rw = 1 To tbl.ListRows.Count
        Set rng = tbl.ListRows(rw).Range
        
        If (ws.Name = "Controle") Then
            Call addRemIcon(ws, rng, rw, 1)
        Else
            Call addEditIcon(ws, rng, rw, 1)
            Call addRemIcon(ws, rng, rw, 2)
        End If
    Next
    Application.ScreenUpdating = True
    
End Sub

Sub deleteIcons()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rw As Integer
    
    Set ws = ActiveSheet
    Set tbl = ws.ListObjects(1)
        
    For rw = 1 To tbl.ListRows.Count
        If (deleteEditIcon(ws, rw) = False And deleteRemIcon(ws, rw) = False) Then Exit For
    Next
    
End Sub

Sub addEditIcon(ws As Worksheet, ByVal rng As Range, ByVal rw As Integer, pos As Integer)
    Dim x As Single, y As Single, size As Single
    Dim pic As Shape
    Dim imgPath As String

    Set rng = rng(1, rng.Columns.Count).offset(0, 1)
    size = rng.Height * 0.65

    x = rng.Left + pos * (IC_SPACE + size) - size
    y = rng.Top + rng.Height / 2 - size / 2
    
    imgPath = ActiveWorkbook.Path & EDIT_PATH
    
    Set pic = ws.Shapes.AddPicture(imgPath, False, True, x, y, size, size)
    pic.Name = "edit" & "_" & rw
    pic.OnAction = "'" & ActiveWorkbook.Name & "'!" & "iniciaAtualiz"
    
End Sub

Sub addRemIcon(ws As Worksheet, ByVal rng As Range, ByVal rw As Integer, pos As Integer)
    Dim x As Single, y As Single, size As Single
    Dim pic As Shape
    Dim imgPath As String
    
    Set rng = rng(1, rng.Columns.Count).offset(0, 1)
    size = rng.Height * 0.65

    x = rng.Left + pos * (IC_SPACE + size) - size
    y = rng.Top + rng.Height / 2 - size / 2
    
    imgPath = ActiveWorkbook.Path & REM_PATH
    
    Set pic = ws.Shapes.AddPicture(imgPath, False, True, x, y, size, size)
    pic.Name = "rem" & "_" & rw
    
    If (ws.Name = "Controle") Then
        pic.OnAction = "'" & ActiveWorkbook.Name & "'!" & "remMov"
    End If
    
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

Sub ajustaIcon(ws As Worksheet, ByVal i As Integer, ByVal nm As String)
    Dim aux As Shape, sh As Shape
    Dim iType As String
    
    Set sh = ws.Shapes(nm)
    
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
End Sub
