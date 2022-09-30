VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} bnmForm 
   Caption         =   "Busca de Produto por Nome"
   ClientHeight    =   7044
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11292
   OleObjectBlob   =   "bnmForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "bnmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim cols_wdth As String
    cols_wdth = "50;100;100;200"
    
    With hList
        .ColumnWidths = cols_wdth
        .AddItem
        .List(0, 0) = "ID"
        .List(0, 1) = "CODIGO DE BARRAS"
        .List(0, 2) = "CODIGO INTERNO"
        .List(0, 3) = "PRODUTO"
    End With
    pList.ColumnWidths = cols_wdth
    
End Sub

Private Sub CommandButton1_Click()
    Dim i As Integer, j As Integer, k As Integer
    Dim pArray() As Variant
    Dim tbl As ListObject
    Set tbl = Sheets("Cadastro").ListObjects(1)
    k = 0
    pList.Clear
    
    pArray = tbl.Range.Value2
    For j = 1 To tbl.HeaderRowRange.Count
        If (pArray(1, j) = "PRODUTO") Then
            Exit For
        End If
    Next
    
    For i = 2 To tbl.Range.Rows.Count
        If (pArray(i, j) Like "*" & UCase(TextBox1) & "*") Then
            With pList
                .AddItem k + 1
                .List(k, 1) = pArray(i, 2)
                .List(k, 2) = pArray(i, 4)
                .List(k, 3) = pArray(i, 5)
            End With
            k = k + 1
        End If
    Next
    
End Sub

Private Sub pList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With mvmForm
        .box1 = pList.List(pList.ListIndex, 2)
        travaCampo .box2, pList.List(pList.ListIndex, 3)
        .box3.SetFocus
    End With
    Unload Me
End Sub
