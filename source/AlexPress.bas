Attribute VB_Name = "AlexPress"
'===============================================================================
'   ������          : AlexPress
'   ������          : 2024.04.07
'   �����           : https://vk.com/elvin_macro/
'                     https://github.com/elvin-nsk/AlexPress
'   �����           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = True

Public Const APP_NAME As String = "AlexPress"

Public Const SEPARATOR As String = "-"
Public Const INFO_SIZE As Double = 3.5 '��
Public Const INFO_SPACE As Double = 1 '��
Public Const INFO_ROUND_DIGITS As Long = 1
Public Const IMPOSITION_PREFIX As String = "����� "
Public Const PLOTTER_MARK_COLOR As String = "CMYK,USER,0,0,0,100"
Public Const PLOTTER_MARK_DIAMETER As Double = 6
Public Const RASTR_RESOLUTION As Long = 300

'� ������� MainLogic.AddMarksAndSeparate
Public Const CUT_COLOR_NAME As String = "CutContour"
Public Const PERFCUT_COLOR_NAME As String = "CutContour Perfcut"
Public Const WHITE_COLOR_NAME As String = "White"
Public Const SEPARATIONS_SHIFT As Double = 25 '��������� �� �������

Public OpenCloseHandler As FilesOpenCloseLogger

'===============================================================================

Sub Start()
    With New MainView
        .Show vbModeless
    End With
End Sub

'===============================================================================

Private Sub TestImposer()

    If ActivePage.Shapes.Count < 2 Then
        VBA.MsgBox "������ ������ 2"
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
        
        Show .ComposedShapesByPage.Count
        MarkComposedShapes .ComposedShapesByPage
    End With
    
    BoostFinish True

End Sub

Private Sub MarkComposedShapes(ByVal ComposedShapes As Collection)
    Dim s As Shape
    Dim i As Long
    Dim a As Double
    For i = 1 To ComposedShapes.Count
        If IsChet(i) Then a = 10 Else a = -10
        For Each s In ComposedShapes(i)
            s.Rotate a
        Next s
    Next i
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

Private Sub TestSticker()
    ActiveDocument.Unit = cdrMillimeter
    With New Sticker
        .SetMainColors _
            FindColor("Summa2", "CutContour"), _
            FindColor("Summa2", "CutContour Perfcut"), _
            FindColor("Summa2", "White")
        .SetDistance 3
        .Separate ActivePage.Shapes.All
    End With
End Sub

Private Sub TestPalette()
    Dim Palette As Palette
    Set Palette = PaletteManager.GetPalette("Summa2")
    Dim Color As Color
    For Each Color In Palette.Colors
        Debug.Print Color.Name
    Next Color
End Sub

Private Sub TestFindColor()
    Debug.Print FindColor("Summa2", "CutContour Perfcut").Tint
End Sub
