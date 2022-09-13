VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    
    Call logout
    Application.DisplayFullScreen = True
    With Sheets("empty")
        .Visible = xlSheetVisible
        .Activate
        .Visible = xlSheetVeryHidden
    End With
    Sheets("Acesso").Activate
    
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    
    Call logout
    Application.DisplayFullScreen = False

End Sub
