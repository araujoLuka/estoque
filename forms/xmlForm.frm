VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} xmlForm 
   Caption         =   "Lista de Produtos do XML"
   ClientHeight    =   6888
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9144
   OleObjectBlob   =   "xmlForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "xmlForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim cols_wdth As String
    
    cols_wdth = "80;80;130;10"

    With hList
        .ColumnWidths = cols_wdth
        .AddItem
        .List(0, 0) = "COD. BARRAS"
        .List(0, 1) = "COD. INTERNO"
        .List(0, 2) = "PRODUTO"
        .List(0, 3) = "QTD"
    End With
    pList.ColumnWidths = cols_wdth
    
    boxU = Range("actv")
    boxData = Date
    boxHora = Time
    
End Sub

Private Sub importBtn_Click()
    
    Me.Hide
    
End Sub

Private Sub cancelBtn_Click()
    
    Unload Me

End Sub

