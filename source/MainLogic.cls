VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MainLogic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

'===============================================================================
' # Declarations

Enum enum_Sides
    LeftSide = 0
    RightSide = 1
    TopSide = 3
    BottomSide = 4
End Enum

Enum enum_Corners
    TopLeftCorner = 0
    TopRightCorner = 1
    BottomLeftCorner = 2
    BottomRightCorner = 3
End Enum

Private Type typeThis
    Cfg As Config
End Type
Private This As typeThis
Attribute This.VB_VarHelpID = -1

'===============================================================================
' # Constructors

Public Function Create( _
                    ByVal View As MainView _
                ) As MainLogic
    Set Create = New MainLogic
    Dim Cfg As Config
    Set Cfg = Config.Bind
    
    With View
        
        .BleedsMin = 0.1
        .BleedsMax = 10000
        .tbBleeds.Value = VBA.CStr(Cfg.Bleeds)
        .cbRound.Value = Cfg.RoundSize
        Select Case Cfg.RoundDec
            Case 0
                .obRound0.Value = True
            Case 1
                .obRound1.Value = True
            Case 2
                .obRound2.Value = True
        End Select
        .cbTrim.Value = Cfg.BitmapTrim
        .tbTrim.Value = Cfg.BitmapTrimSize
        .cbFlatten.Value = Cfg.BitmapFlatten
        .tbPages.Value = Cfg.PagesToAdd
        .PressSheetWidth = Cfg.PressSheetWidth
        .PressSheetHeight = Cfg.PressSheetHeight
        .PressSheetSpaces = Cfg.PressSheetSpaces
        .OptionWithBacks = Cfg.OptionWithBacks
        .OptionMarks = Cfg.OptionMarks
        .OptionNoBacks = Cfg.OptionNoBacks
    
    End With
    
    Create.Inject Cfg
End Function

Friend Sub Inject( _
               ByVal Cfg As Config _
           )
    With This
        Set .Cfg = Cfg
    End With
End Sub

'===============================================================================
' # Logic

Public Function AddBleeds(ByVal View As MainView)

    If AlexPress.RELEASE Then On Error GoTo Catch
    
    FillCfgFromView View
    
    BoostStart "Добавить припуски", AlexPress.RELEASE
        
    SetBleeds This.Cfg
    
Finally:
    BoostFinish
    Exit Function

Catch:
    MsgBox "Ошибка: " & Err.Description, vbCritical
    Resume Finally
End Function

Public Function AddPages(ByVal View As MainView)

    If AlexPress.RELEASE Then On Error GoTo Catch
    
    FillCfgFromView View
    
    BoostStart "Добавить страницы", AlexPress.RELEASE
        
    AddPagesRoutine This.Cfg
    
Finally:
    BoostFinish
    Exit Function

Catch:
    MsgBox "Ошибка: " & Err.Description, vbCritical
    Resume Finally
End Function

Public Function MakeImposition( _
                    ByVal View As MainView, _
                    ByVal Motifs As MotifsCollector _
                )

    If AlexPress.RELEASE Then On Error GoTo Catch
    
    FillCfgFromView View
    
    ActiveDocument.Unit = cdrMillimeter
    BoostStart "Подгонка размеров", AlexPress.RELEASE
        ResizeMotifs Motifs.AsCollection
    BoostFinish
    
    Dim TempFiles As Collection
    Set TempFiles = Motifs.ExportToTemp
    Dim Doc As Document
    Set Doc = CreateDocument
    Doc.Name = "Спуск " _
             & This.Cfg.PressSheetWidth _
             & "x" _
             & This.Cfg.PressSheetHeight
    Doc.Unit = cdrMillimeter
    Doc.MasterPage.SizeWidth = This.Cfg.PressSheetWidth
    Doc.MasterPage.SizeHeight = This.Cfg.PressSheetHeight
    Doc.Activate
    
    If RotatePageToMaxFit( _
           SpaceBox(Doc.MasterPage.BoundingBox, -This.Cfg.PressSheetSpaces), _
           Motifs.Item(1).SurfaceA.Content.BoundingBox _
       ) Then
        SwapOrientation Doc.MasterPage
    End If
    
    BoostStart "Спуск", AlexPress.RELEASE
        
    With MotifsImposer.CreateFromImport( _
             Motifs.AsCollection, _
             This.Cfg.OptionWithBacks _
         )
        .Bleeds = This.Cfg.Bleeds
        .EdgeMarginBottom = This.Cfg.PressSheetSpaces
        .EdgeMarginLeft = This.Cfg.PressSheetSpaces
        .EdgeMarginRight = This.Cfg.PressSheetSpaces
        .EdgeMarginTop = This.Cfg.PressSheetSpaces
        .FillLastPage = True
        .PlaceCropMarks = This.Cfg.OptionMarks
        .PlaceWhiteBackground = True
        .ImposeAutoAddPages
    End With
    
    DeleteFiles TempFiles
    
Finally:
    BoostFinish
    Exit Function

Catch:
    MsgBox "Ошибка: " & Err.Description, vbCritical
    Resume Finally
End Function

Public Sub NumberSetter()

    If AlexPress.RELEASE Then On Error GoTo Catch
        
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
    
    BoostStart "Расстановка номеров", AlexPress.RELEASE
    
    SetNumbers Shapes, Prefix, StartingNumber
    
Finally:
    BoostFinish
    Exit Sub

Catch:
    MsgBox "Ошибка: " & Err.Description, vbCritical
    Resume Finally
End Sub

'===============================================================================
' # Public

Public Property Get RotatePageToMaxFit( _
                        ByVal ImpositionArea As Rect, _
                        ByVal Box As Rect _
                    ) As Boolean
    Dim NumberAsIs As Long
    Dim NumberIfRotate As Long
    NumberAsIs = NumberToFitArea(Box, ImpositionArea)
    NumberIfRotate = NumberToFitArea(Box, GetRotatedRect(ImpositionArea))
    RotatePageToMaxFit = NumberIfRotate > NumberAsIs
End Property

Public Function Dispose(ByVal View As MainView)
    FillCfgFromView View
End Function

Public Property Get Self() As MainLogic
    Set Self = Me
End Property

'===============================================================================
' # Helpers

Private Sub FillCfgFromView(ByVal View As MainView)
    With View
        
        This.Cfg.Bleeds = VBA.CDbl(.tbBleeds.Value)
        This.Cfg.RoundSize = .cbRound.Value
        Select Case True
            Case .obRound0.Value = True
                This.Cfg.RoundDec = 0
            Case .obRound1.Value = True
                This.Cfg.RoundDec = 1
            Case .obRound2.Value = True
                This.Cfg.RoundDec = 2
        End Select
        This.Cfg.BitmapTrim = .cbTrim.Value
        This.Cfg.BitmapTrimSize = .tbTrim.Value
        This.Cfg.BitmapFlatten = .cbFlatten.Value
        This.Cfg.PagesToAdd = .tbPages.Value
        This.Cfg.PressSheetWidth = .PressSheetWidth
        This.Cfg.PressSheetHeight = .PressSheetHeight
        This.Cfg.PressSheetSpaces = .PressSheetSpaces
        This.Cfg.OptionWithBacks = .OptionWithBacks
        This.Cfg.OptionMarks = .OptionMarks
        This.Cfg.OptionNoBacks = .OptionNoBacks
    
    End With
End Sub

Private Function SetBleeds(ByVal Cfg As Config)
    
    Dim SrcShape As Shape, Bleeds As Shape, Final As Shape
    Dim Range As New ShapeRange
    Dim Width As Double, Height As Double, Name As String
    
    Set SrcShape = ActiveSelectionRange.FirstShape
    
    'если округляем размер
    If Cfg.RoundSize Then
        Width = VBA.Round(SrcShape.SizeWidth, Cfg.RoundDec)
        Height = VBA.Round(SrcShape.SizeHeight, Cfg.RoundDec)
    Else
        Width = SrcShape.SizeWidth
        Height = SrcShape.SizeHeight
    End If
    
    'если подчищаем битмап
    If SrcShape.Type = cdrBitmapShape Then
        If Cfg.BitmapTrim Then
            ShrinkBitmap SrcShape, Cfg.BitmapTrimSize
        End If
    End If
    
    SrcShape.SetSize Width, Height
    
    Set Bleeds = CreateBleeds(SrcShape, Cfg.Bleeds)
    
    'если растрируем всё обратно в битмап
    If SrcShape.Type = cdrBitmapShape And Cfg.BitmapFlatten Then
        Name = SrcShape.Name
        Set Final = Flatten(SrcShape, Bleeds)
        Final.Name = Name
    Else 'а нет - так группируем, обзываем
        Bleeds.Name = "припуски"
        Range.Add Bleeds
        Range.Add SrcShape
        Set Final = Range.Group
        If SrcShape.Name = "" Then
            Final.Name = "группа - растр с припусками"
        Else
            Final.Name = SrcShape.Name & " (группа с припусками)"
        End If
    End If
    
    Final.CreateSelection

End Function

Private Sub ShrinkBitmap( _
                ByRef ioBitmapShape As Shape, _
                ByVal Pixels As Long _
            )
    
    Dim Crop As Shape
    Dim PxW As Double, PxH As Double
    Dim SizeW As Double, SizeH As Double, AngleMult As Long
    Dim SaveUnit As cdrUnit, SavePoint As cdrReferencePoint
    
    If ioBitmapShape.Type <> cdrBitmapShape Then Exit Sub
    If Pixels < 1 Then Exit Sub
    
    'save
    SaveUnit = ActiveDocument.Unit
    SavePoint = ActiveDocument.ReferencePoint
    
    ActiveDocument.Unit = cdrInch
    ActiveDocument.ReferencePoint = cdrCenter
    With ioBitmapShape
        SizeW = .SizeWidth
        SizeH = .SizeHeight
        AngleMult = .RotationAngle \ 90
        .ClearTransformations
        .RotationAngle = AngleMult * 90
        .SetSize SizeW, SizeH
        PxW = 1 / .Bitmap.ResolutionX
        PxH = 1 / .Bitmap.ResolutionY
        Set Crop = _
            .Layer.CreateRectangle( _
                ioBitmapShape.LeftX + PxW * Pixels, _
                .TopY - PxH * Pixels, _
                .RightX - PxW * Pixels, _
                .BottomY + PxH * Pixels _
            )
    End With
    Set ioBitmapShape = TrimBitmap(ioBitmapShape, Crop, False)
    
    'restore
    ActiveDocument.Unit = SaveUnit
    ActiveDocument.ReferencePoint = SavePoint

End Sub

Private Function CreateBleeds( _
                     ByVal BitmapShape As Shape, _
                     ByVal Bleed As Double _
                 ) As Shape
    
    Dim Shapes As New ShapeRange
    
    On Error Resume Next
    
    With Shapes
        .Add CreateSideBleed(BitmapShape, Bleed, LeftSide)
        .Add CreateSideBleed(BitmapShape, Bleed, RightSide)
        .Add CreateSideBleed(BitmapShape, Bleed, TopSide)
        .Add CreateSideBleed(BitmapShape, Bleed, BottomSide)
        
        .Add CreateCornerBleed(BitmapShape, Bleed, TopLeftCorner)
        .Add CreateCornerBleed(BitmapShape, Bleed, TopRightCorner)
        .Add CreateCornerBleed(BitmapShape, Bleed, BottomLeftCorner)
        .Add CreateCornerBleed(BitmapShape, Bleed, BottomRightCorner)
    End With
    
    On Error GoTo 0
    
    Set CreateBleeds = Shapes.Group

End Function

Private Function Flatten( _
                     ByVal SourceBitmap As Shape, _
                     ByVal BleedsGroup As Shape _
                 ) As Shape
    Dim Range As New ShapeRange
    Dim Width#, Height#
    Range.Add SourceBitmap
    Range.Add BleedsGroup
    Width = Range.SizeWidth
    Height = Range.SizeHeight
    Range.SetPixelAlignedRendering True
    With SourceBitmap.Bitmap
        If .ResolutionX <> .ResolutionY Then
            Range.SizeHeight = _
                Range.SizeHeight * .ResolutionY / .ResolutionX
        End If
        Set Flatten = _
            Range.ConvertToBitmapEx( _
                .Mode, , .Transparent, .ResolutionX, _
                cdrNoAntiAliasing, False _
            )
    End With
    Flatten.SetSize Width, Height
End Function

Private Function CreateSideBleed( _
                     ByVal BitmapShape As Shape, _
                     ByVal Bleed As Double, _
                     ByVal Side As enum_Sides _
                 ) As Shape
    
    Dim LeftAdd As Double, RightAdd As Double
    Dim TopAdd As Double, BottomAdd As Double
    Dim ShiftX As Double, ShiftY As Double
    Dim Flip As cdrFlipAxes
    
    Select Case Side
        Case LeftSide
            RightAdd = -(BitmapShape.SizeWidth - Bleed)
            Flip = cdrFlipHorizontal
            ShiftX = -Bleed
        Case RightSide
            LeftAdd = BitmapShape.SizeWidth - Bleed
            Flip = cdrFlipHorizontal
            ShiftX = Bleed
        Case TopSide
            BottomAdd = BitmapShape.SizeHeight - Bleed
            Flip = cdrFlipVertical
            ShiftY = Bleed
        Case BottomSide
            TopAdd = -(BitmapShape.SizeHeight - Bleed)
            Flip = cdrFlipVertical
            ShiftY = -Bleed
    End Select
    
    Set CreateSideBleed = _
        CropTool( _
            BitmapShape.Duplicate, BitmapShape.LeftX + LeftAdd, _
            BitmapShape.TopY + TopAdd, _
            BitmapShape.RightX + RightAdd, _
            BitmapShape.BottomY + BottomAdd _
        ).FirstShape
    If CreateSideBleed Is Nothing Then Exit Function
    CreateSideBleed.Flip Flip
    CreateSideBleed.Move ShiftX, ShiftY
    CreateSideBleed.Name = "боковой припуск"

End Function

Private Function CreateCornerBleed( _
                     ByVal BitmapShape As Shape, _
                     ByVal Bleed As Double, _
                     ByVal Corner As enum_Corners _
                 ) As Shape
    
    Dim LeftAdd As Double, RightAdd As Double
    Dim TopAdd As Double, BottomAdd As Double
    Dim ShiftX As Double, ShiftY As Double
    
    Select Case Corner
        Case TopLeftCorner
            RightAdd = -(BitmapShape.SizeWidth - Bleed)
            BottomAdd = BitmapShape.SizeHeight - Bleed
            ShiftX = -Bleed
            ShiftY = Bleed
        Case TopRightCorner
            LeftAdd = BitmapShape.SizeWidth - Bleed
            BottomAdd = BitmapShape.SizeHeight - Bleed
            ShiftX = Bleed
            ShiftY = Bleed
        Case BottomLeftCorner
            RightAdd = -(BitmapShape.SizeWidth - Bleed)
            TopAdd = -(BitmapShape.SizeHeight - Bleed)
            ShiftX = -Bleed
            ShiftY = -Bleed
        Case BottomRightCorner
            LeftAdd = BitmapShape.SizeWidth - Bleed
            TopAdd = -(BitmapShape.SizeHeight - Bleed)
            ShiftX = Bleed
            ShiftY = -Bleed
    End Select
    Set CreateCornerBleed = _
        CropTool( _
            BitmapShape.Duplicate, BitmapShape.LeftX + LeftAdd, _
            BitmapShape.TopY + TopAdd, _
            BitmapShape.RightX + RightAdd, _
            BitmapShape.BottomY + BottomAdd _
        ).FirstShape
    If CreateCornerBleed Is Nothing Then Exit Function
    CreateCornerBleed.Flip cdrFlipBoth
    CreateCornerBleed.Move ShiftX, ShiftY
    CreateCornerBleed.Name = "угловой припуск"
    
End Function

'----------------------------------------------------------------

Private Function AddPagesRoutine(ByVal Cfg As Config)
    Dim Index As Long
    Dim AfterPageIndex As Long
    AfterPageIndex = ActivePage.Index
    For Index = 1 To Cfg.PagesToAdd
        AddPage AfterPageIndex, Cfg
    Next Index
    ActiveDocument.Pages(AfterPageIndex).Activate
End Function

Private Function AddPage(ByVal AfterPageIndex As Long, ByVal Cfg As Config) As Page
    ActiveDocument.InsertPages 1, False, AfterPageIndex
    Dim Page As Page
    Set Page = ActiveDocument.Pages(AfterPageIndex)
    Page.Activate
    Dim ShapesToCopy As ShapeRange
    Dim TargetLayer As Layer
    Dim Layer As Layer
    For Each Layer In Page.Layers
        If Not Layer.IsSpecialLayer Then
            Set ShapesToCopy = GetShapesOnPage(Layer.Shapes.All, Page)
            If ShapesToCopy.Count > 0 Then
                ShapesToCopy.CopyToLayer Page.Next.Layers.Find(Layer.Name)
            End If
        End If
        Page.Activate
    Next Layer
End Function

Private Property Get GetShapesOnPage( _
                         ByVal Shapes As ShapeRange, _
                         ByVal Page As Page _
                     ) As ShapeRange
    Set GetShapesOnPage = CreateShapeRange
    Dim Shape As Shape
    For Each Shape In Shapes.ReverseRange
        If Not Shape.BoundingBox.Intersect(Page.BoundingBox).IsEmpty Then
            GetShapesOnPage.Add Shape
        End If
    Next Shape
End Property

Private Sub ResizeMotifs(ByVal Motifs As Collection)
    Dim Width As Double, Height As Double
    Dim SrcShape As Shape
    Set SrcShape = Motifs(1).SurfaceA.Content
    With This
        If .Cfg.RoundSize Then
            Width = VBA.Round(SrcShape.SizeWidth, .Cfg.RoundDec)
            Height = VBA.Round(SrcShape.SizeHeight, .Cfg.RoundDec)
        Else
            Width = SrcShape.SizeWidth
            Height = SrcShape.SizeHeight
        End If
        Dim Motif As Motif
        For Each Motif In Motifs
            Motif.SurfaceA.Content.SetSize Width, Height
            If Motif.HasSurfaceB Then
                Motif.SurfaceB.Content.SetSize Width, Height
            End If
        Next Motif
    End With
End Sub

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

Private Sub DeleteFiles(ByVal Files As Collection)
    Dim File As IFileSpec
    For Each File In Files
        VBA.Kill File
    Next File
End Sub
