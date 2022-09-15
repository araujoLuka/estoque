VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim rw As Integer
    Dim arr As Variant
    
    Application.ScreenUpdating = False
    
    With ActiveSheet.ListObjects(1).Range
        .Borders.LineStyle = xlNone
        .Interior.ColorIndex = xlNone
    End With
    
    If (Not Target.ListObject Is Nothing) Then
        rw = Target.Row - Target.ListObject.HeaderRowRange.Row
        
        If (rw <= 0) Then Exit Sub
        With Target.ListObject.ListRows(rw)
            '.Range.Interior.Color = RGB(245, 245, 245)
            With .Range
                .Interior.Color = RGB(230, 230, 230)
                With .Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Color = -11489280
                    .Weight = xlThin
                End With
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Color = -11489280
                    .Weight = xlThin
                End With
                With .Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Color = -11489280
                    .Weight = xlMedium
                End With
            End With
        End With
    End If
    
    Application.ScreenUpdating = True

End Sub
