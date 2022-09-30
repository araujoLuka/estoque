VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const BACKUP_DIR As String = "\others\"
Private Const BACKUP_FILENAME As String = BACKUP_DIR & "backup.xlsx"

Private Sub Workbook_Open()
        
    Application.ScreenUpdating = False
    Call logout
    With Sheets("empty")
        .Visible = xlSheetVisible
        .Activate
    End With
    Application.DisplayFullScreen = True
    Sheets("Acesso").Activate
    Sheets("empty").Visible = xlSheetVeryHidden
    Application.ScreenUpdating = True
    
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
        
    Call logout
    With Sheets("empty")
        .Visible = xlSheetVisible
        .Activate
    End With
    Application.DisplayFullScreen = False
    Save

End Sub
