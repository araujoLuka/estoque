Attribute VB_Name = "tblScripts"
Option Explicit

Function insereRow(tbl As ListObject) As Range
    Dim rng As Range
    
    If (tbl.Range(2, 1) <> "") Then
        Set rng = tbl.ListRows.Add().Range
        tbl.ListRows(1).Range.EntireRow.Copy
        rng.EntireRow.PasteSpecial xlPasteFormats
        rng.EntireRow.Hidden = False
        Application.CutCopyMode = False
    Else
        If (tbl.ListRows.Count = 0) Then
            Set rng = tbl.ListRows.Add().Range
        Else
            Set rng = tbl.ListRows(1).Range
        End If
    End If
    
    Set insereRow = rng

End Function

Sub deleteRow(ws As Worksheet, rng As Range)
    Dim tbl As ListObject
    Dim aux As Range
    
    Set tbl = ws.ListObjects(1)
          
    rng.EntireRow.Delete
    
    If (tbl.ListRows.Count <> 0) Then
        Set aux = tbl.ListRows(tbl.ListRows.Count).Range.Offset(1)
        
        While (aux.EntireRow.Hidden = True)
            aux.EntireRow.Hidden = False
            aux.EntireRow.Hidden = True
            Set aux = aux.Offset(1)
        Wend
    Else
        Set aux = tbl.HeaderRowRange.Offset(2)
    End If
        
    aux.EntireRow.Hidden = True
End Sub

Sub sortTbl(cTabble As ListObject)
    Dim c1 As Integer, c2 As Integer
    Dim i As Integer
    Dim arr As Variant
    Dim s As String

    arr = cTabble.Range.Value2
    For i = 1 To UBound(arr, 2)
        s = arr(1, i)
        If (s = "TIPO") Then
            c1 = i
        ElseIf (s = "PRODUTO") Then
            c2 = i
        End If
        If (c1 <> 0 And c2 <> 0) Then Exit For
    Next
    
    If (i > UBound(arr, 2)) Then
        If (c1 = 0) Then c1 = c2
    End If
    
    If (cTabble.ListRows.Count > 0) Then
        cTabble.Range.Sort Key1:=cTabble.ListColumns(c1), _
                           Order1:=xlAscending, _
                           Header:=xlYes, _
                           Key2:=cTabble.ListColumns(c2), _
                           Order2:=xlAscending, _
                           Header:=xlYes
    End If
End Sub
