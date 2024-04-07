VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisMacroStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub GlobalMacroStorage_Quit()
    On Error GoTo Catch
    
    Set OpenCloseHandler = Nothing
    
Finally:
    Exit Sub

Catch:
    Debug.Print "Error: " & VBA.Err.Description
    Resume Finally
End Sub

Private Sub GlobalMacroStorage_Start()
    On Error GoTo Catch
    
    Dim Cfg As Config
    Set Cfg = Config.Bind_
    If Not Cfg.WorkLogsEnabled Then GoTo Finally
    If Not FSO.FolderExists(Cfg.WorkLogsPath) Then
        VBA.MsgBox _
            "Не найден путь для отчётов: " & Cfg.WorkLogsPath, _
            vbExclamation
        Exit Sub
    End If
    Set OpenCloseHandler = FilesOpenCloseLogger.New_(Cfg.WorkLogsPath)
    
Finally:
    Set Cfg = Nothing
    Exit Sub

Catch:
    Debug.Print "Error: " & VBA.Err.Description
    Resume Finally
End Sub
