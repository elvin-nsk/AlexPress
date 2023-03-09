Attribute VB_Name = "AlexPress"
'===============================================================================
'   Макрос          : AlexPress
'   Версия          : 2023.03.10
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

Sub NumberSetter()

    If RELEASE Then On Error GoTo Catch
        
    Dim Shapes As ShapeRange
    With InitData.GetShapes(ErrNoSelection:="Выделите текстовые объекты")
        If .IsError Then Exit Sub
        Set Shapes = .Shapes.Shapes.FindShapes(Type:=cdrTextShape)
    End With
    If Shapes.Count = 0 Then
        Throw "Среди выделенных объектов нет ни одного текстового"
    End If
    
    Dim Prefix As String
    Dim StartingNumber As Long
    With New NumberSetterView
        .Show
        If .IsCancel Then Exit Sub
        Prefix = .Prefix
        StartingNumber = .StartingNumber
    End With
    
    BoostStart "NumberSetter", RELEASE
    
    SetNumbers Shapes, Prefix, StartingNumber
    
Finally:
    BoostFinish
    Exit Sub

Catch:
    MsgBox "Ошибка: " & Err.Description, vbCritical
    Resume Finally
End Sub

'===============================================================================

Private Sub SetNumbers( _
                ByVal Shapes As ShapeRange, _
                ByVal Prefix As String, _
                ByVal StartingNumber As Long _
            )
    Dim NewShapes As ShapeRange
    Dim Page As Page
    For Each Page In ActiveDocument.Pages
        Set NewShapes = Shapes.Duplicate
        NewShapes.MoveToLayer Page.ActiveLayer
        SetNumberingText NewShapes, Prefix, StartingNumber
    Next Page
    Shapes.Delete
End Sub

Private Sub SetNumberingText( _
                ByVal Shapes As ShapeRange, _
                ByVal Prefix As String, _
                ByRef ioNextStartingNumber As Long _
            )
    Dim Shape As Shape
    For Each Shape In Shapes
        Shape.Text.Story.Text = Prefix & SEPARATOR & ioNextStartingNumber
        ioNextStartingNumber = ioNextStartingNumber + 1
    Next Shape
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
        .Bleeds = 2
        .PlaceCropMarks = True
        .FillLastPage = True
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
