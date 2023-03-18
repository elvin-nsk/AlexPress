Attribute VB_Name = "AlexPress"
'===============================================================================
'   Макрос          : AlexPress
'   Версия          : 2023.03.18
'   Сайты           : https://vk.com/elvin_macro/
'                     https://github.com/elvin-nsk/AlexPress
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = True

Public Const APP_NAME As String = "AlexPress"
Public Const SEPARATOR As String = "-"

'===============================================================================

Sub Start()
    With New MainView
        .Show vbModeless
    End With
End Sub

'===============================================================================

Private Sub TestImposer()

    If ActivePage.Shapes.Count < 2 Then
        VBA.MsgBox "Шейпов меньше 2"
        Exit Sub
    End If
    
    Dim Motifs As New MotifsCollector
    Dim Surface As Surface
    Dim Index As Long
    For Index = 2 To ActivePage.Shapes.Count Step 2
        With New Motif
            Set Surface = New Surface
            Set Surface.Content = ActivePage.Shapes(Index - 1)
            Set .SurfaceA = Surface
            Set Surface = New Surface
            Set Surface.Content = ActivePage.Shapes(Index)
            Set .SurfaceB = Surface
            .Quantity = Index * 2
            Motifs.Add .Self
        End With
    Next Index
    
    'Show Motifs.Count
    Motifs.ExportToTemp
    
    Dim Doc As Document
    Set Doc = CreateDocument
    Doc.Name = "imposer test"
    Doc.Unit = cdrMillimeter
    Doc.MasterPage.SetSize 300, 200
    
    BoostStart "TestImposer", True
   
    With MotifsImposer.CreateFromImport(Motifs.AsCollection, True)
        '.ReferencePoint = cdrTopLeft
        .Bleeds = 2
        '.PlaceCropMarks = True
        .PlaceWhiteBackground = True
        .FillLastPage = True
        .EdgeMarginBottom = 30
        .EdgeMarginLeft = 30
        .EdgeMarginRight = 30
        .EdgeMarginTop = 30
        
        .ImposeAutoAddPages
    End With
    
    BoostFinish True

End Sub

Private Sub TestMarks()
    BoostStart "TestMarks", False
    With MarksSetter.Create(ActivePage.Shapes.All)
        .Bleeds = 2
        .Size = 3
        .PlaceWhiteUndermark = True
        .SetMarksOnNewLayer
    End With
    BoostFinish True
End Sub

Private Sub TestOther()
    'ActiveDocument.Unit = cdrMillimeter
    Show ActiveSelectionRange.FirstShape.BoundingBox.Width
End Sub
