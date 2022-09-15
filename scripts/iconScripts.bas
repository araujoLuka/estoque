Attribute VB_Name = "iconScripts"
Option Explicit

Private Const EDIT_PATH As String = ".\resources\edit_icon1.png"
Private Const REM_PATH As String = ".\resources\rem_icon2.png"

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
            .Left = rng.Left + (rng.Offset(0, 1).Left - rng.Left) / 2 - sh.Width / 3
            .Top = rng.Top + (rng.Offset(1, 0).Top - rng.Top) / 2 - sh.Height / 2
        End With
    Next
    
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

Sub addIcons()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim rw As Integer
    Dim size As Single
    Dim space As Integer
    
    Set ws = ActiveSheet
    Set tbl = ws.ListObjects(1)
    
    space = 15
    
    For rw = 1 To tbl.ListRows.Count
        Set rng = tbl.ListRows(rw).Range(1, tbl.ListColumns.Count).Offset(0, 1)
        size = rng.Height * 0.65
        
        Call addEditIcon(ws, rng, rw, size, space)
        Call addRemIcon(ws, rng, rw, size, space)
    Next
End Sub

Sub addEditIcon(ws As Worksheet, ByVal rng As Range, ByVal rw As Integer, _
                    ByVal size As Single, ByVal space As Single)
    Dim x As Single, y As Single
    Dim pic As Shape
    Dim imgPath As String

    x = rng.Left + space
    y = rng.Top + rng.Height / 2 - size / 2
    
    imgPath = ActiveWorkbook.Path & EDIT_PATH
    
    Set pic = ws.Shapes.AddPicture(imgPath, False, True, x, y, size, size)
    pic.Name = "edit" & "_" & rw
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

Sub addRemIcon(ws As Worksheet, ByVal rng As Range, ByVal rw As Integer, _
                    ByVal size As Single, ByVal space As Single)
    Dim x As Single, y As Single
    Dim pic As Shape
    Dim imgPath As String
    
    x = rng.Left + space + space + size
    y = rng.Top + rng.Height / 2 - size / 2
    
    imgPath = ActiveWorkbook.Path & REM_PATH
    
    Set pic = ws.Shapes.AddPicture(imgPath, False, True, x, y, size, size)
    pic.Name = "rem" & "_" & rw
    
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

Sub invSplitRemIcon(ws As Worksheet, ByVal i As Integer, sh As Shape)
    Dim nm As String
    Dim aux As Shape
    
    For i = i To ws.ListObjects(1).ListRows.Count - 1
        nm = "rem_" & i + 1
        
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
