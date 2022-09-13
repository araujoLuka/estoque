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
    Dim i As Integer
    
    Set tbl = ws.ListObjects(1)
    Set aux = tbl.ListRows(tbl.ListRows.Count).Range
        
    rng.EntireRow.Delete
    
    aux.Offset(1).EntireRow.Hidden = True
End Sub

Sub sortCad(cTabble As ListObject)

    cTabble.Range.Sort Key1:=cTabble.ListColumns(3), _
                       Order1:=xlAscending, _
                       Header:=xlYes, _
                       Key2:=cTabble.ListColumns(5), _
                       Order2:=xlAscending, _
                       Header:=xlYes
End Sub
