VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    If (ActiveSheet.Name <> Me.Name) Then Exit Sub
    
    If (Target.Count > 1) Then Exit Sub
    
    If (Target.End(xlUp) = "LIMITE" And Not Target.Formula Like "=*") Then
        MsgBox "Para alterar o limite de estoque acesse a aba 'Cadastro'"
        Application.Undo
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim rw As Integer
    Dim arr As Variant
        
    If (Not Application.CutCopyMode = False Or (Selection.Address <> Target.Address)) Then Exit Sub
    
    With ActiveSheet.ListObjects(1).Range
        .Borders.LineStyle = xlNone
        .Interior.ColorIndex = xlNone
    End With
    
    If (Not Target.ListObject Is Nothing) Then
        Call highlightSelection(Target)
    End If
End Sub

