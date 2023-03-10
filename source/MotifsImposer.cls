VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MotifsImposer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
'   ??????          : MotifsImposer
'   ??????          : 2023.02.25
'   ?????           : elvin-nsk (me@elvin.nsk.ru)
'   ????            : https://github.com/elvin-nsk/LowCoupledFromCore
'   ??????????      :
'   ???????????     : LibCore, Composer, IComposerElement,
'                     Motif, Surface, IPoint, MarksSetter
'===============================================================================

'@ModuleDescription("MotifsImposer.")

'@PredeclaredId
Option Explicit

'===============================================================================
' # Declarations

Private Type typeThis
    Bleeds As Double
    ComposedElements As Collection
    CropMarksLength As Double
    EdgeMarginBottom As Double
    EdgeMarginLeft As Double
    EdgeMarginRight As Double
    EdgeMarginTop As Double
    Elements As Collection
    FillLastPage As Boolean
    HorizontalSpace As Double
    MaxPlacesInHeight As Long
    MaxPlacesInWidth As Long
    PlaceCropMarks As Boolean
    ReferencePoint As cdrReferencePoint
    RemainingElements As Collection
    TwoSided As Boolean
    VerticalSpace As Double
End Type
Private This As typeThis

Private Const SIDE_A_MARK As String = "A"
Private Const SIDE_B_MARK As String = "B"

Public Event OnImportFile()
Public Event OnPageImposed()

'===============================================================================
' # Constructors

Public Function CreateFromImport( _
                    ByRef Motifs As Variant, _
                    Optional ByVal TwoSided As Boolean = False _
                ) As MotifsImposer
    Set CreateFromImport = New MotifsImposer
    CreateFromImport.Inject Motifs, TwoSided
End Function

Friend Sub Inject( _
               ByRef Motifs As Variant, _
               ByVal TwoSided As Boolean _
           )
    With This
        .CropMarksLength = 3
        .FillLastPage = False
        .ReferencePoint = cdrCenter
        .TwoSided = TwoSided
        Set .Elements = ImportAsMarkedElements(Motifs)
    End With
End Sub

'===============================================================================
' # Public

Public Property Get Bleeds() As Double
    Bleeds = This.Bleeds
End Property
Public Property Let Bleeds(RHS As Double)
    This.Bleeds = RHS
End Property

Public Property Get ComposedElements() As Collection
    Set ComposedElements = This.ComposedElements
End Property

Public Property Get CropMarksLength() As Double
    CropMarksLength = This.CropMarksLength
End Property
Public Property Let CropMarksLength(RHS As Double)
    This.CropMarksLength = RHS
End Property

Public Property Get EdgeMarginBottom() As Double
    EdgeMarginBottom = This.EdgeMarginBottom
End Property
Public Property Let EdgeMarginBottom(RHS As Double)
    This.EdgeMarginBottom = RHS
End Property

Public Property Get EdgeMarginLeft() As Double
    EdgeMarginLeft = This.EdgeMarginLeft
End Property
Public Property Let EdgeMarginLeft(RHS As Double)
    This.EdgeMarginLeft = RHS
End Property

Public Property Get EdgeMarginRight() As Double
    EdgeMarginRight = This.EdgeMarginRight
End Property
Public Property Let EdgeMarginRight(RHS As Double)
    This.EdgeMarginRight = RHS
End Property

Public Property Get EdgeMarginTop() As Double
    EdgeMarginTop = This.EdgeMarginTop
End Property
Public Property Let EdgeMarginTop(RHS As Double)
    This.EdgeMarginTop = RHS
End Property

Public Property Get FillLastPage() As Boolean
    FillLastPage = This.FillLastPage
End Property
Public Property Let FillLastPage(RHS As Boolean)
    This.FillLastPage = RHS
End Property

Public Property Get HorizontalSpace() As Double
    HorizontalSpace = This.HorizontalSpace
End Property
Public Property Let HorizontalSpace(RHS As Double)
    This.HorizontalSpace = RHS
End Property

Public Sub ImposeOnCurrentPageOnly()
    Impose False
End Sub

Public Sub ImposeAutoAddPages()
    Impose True
End Sub

Public Property Get MaxPlacesInHeight() As Long
    MaxPlacesInHeight = This.MaxPlacesInHeight
End Property
Public Property Let MaxPlacesInHeight(RHS As Long)
    This.MaxPlacesInHeight = RHS
End Property

Public Property Get MaxPlacesInWidth() As Long
    MaxPlacesInWidth = This.MaxPlacesInWidth
End Property
Public Property Let MaxPlacesInWidth(RHS As Long)
    This.MaxPlacesInWidth = RHS
End Property

Public Property Get PlaceCropMarks() As Boolean
    PlaceCropMarks = This.PlaceCropMarks
End Property
Public Property Let PlaceCropMarks(RHS As Boolean)
    This.PlaceCropMarks = RHS
End Property

Public Property Get ReferencePoint() As cdrReferencePoint
    ReferencePoint = This.ReferencePoint
End Property
Public Property Let ReferencePoint(RHS As cdrReferencePoint)
    This.ReferencePoint = RHS
End Property

Public Property Get RemainingElements() As Collection
    Set RemainingElements = This.RemainingElements
End Property

Public Property Get Self() As MotifsImposer
    Set Self = Me
End Property

Public Property Get VerticalSpace() As Double
    VerticalSpace = This.VerticalSpace
End Property
Public Property Let VerticalSpace(RHS As Double)
    This.VerticalSpace = RHS
End Property

'===============================================================================
' # Logic

Private Sub Impose( _
                Optional ByVal AutoAddPages As Boolean _
            )
    
    With This
    
        Dim Elements As Collection
        Set Elements = .Elements
        Dim Impositions As Collection
        Set Impositions = New Collection
        Dim ImpositionIndex As Long
        Dim LastPivot As IPoint
        Dim LastCount As Long
        Dim LastIteration As Boolean
        Dim ComposedShapes As ShapeRange
        
        Dim ElementsComposer As Composer
    
        Do
            ImpositionIndex = ImpositionIndex + 1
            Set ElementsComposer = _
                Composer.CreateAndCompose( _
                    Elements:=Elements, _
                    StartingPoint:=FreePoint.Create(-20000, 20000), _
                    MaxPlacesInWidth:=MaxPlacesInWidth, _
                    MaxPlacesInHeight:=MaxPlacesInHeight, _
                    MaxWidth:= _
                        ActivePage.SizeWidth _
                     - .EdgeMarginLeft _
                     - .EdgeMarginRight, _
                    MaxHeight:= _
                        ActivePage.SizeHeight _
                     - .EdgeMarginTop _
                     - .EdgeMarginBottom, _
                    HorizontalSpace:=.HorizontalSpace, _
                    VerticalSpace:=.VerticalSpace _
                )
            
            Set ComposedShapes = _
                ElementsToShapes(ElementsComposer.ComposedElements)
            
            Set LastPivot = AlignImposition(ComposedShapes, LastPivot)
            LastCount = ElementsComposer.ComposedElements.Count
            
            RaiseEvent OnPageImposed
                          
            If Not AutoAddPages Then Exit Do
            
            Impositions.Add ComposedShapes
            Set Elements = ElementsComposer.RemainingElements
            
            If Elements.Count < LastCount _
           And Elements.Count > 0 _
           And .FillLastPage Then
                If .FillLastPage Then
                    AddRandomElements _
                        Elements, _
                        LastCount - ElementsComposer.RemainingElements.Count
                End If
            End If
            
        Loop While Elements.Count > 0
        
        If AutoAddPages Then
            ProcessImpositions Impositions
        End If
    
    End With
    
End Sub

'===============================================================================
' # Helpers

Private Function AlignImposition( _
                ByVal ComposedShapes As ShapeRange, _
                ByVal Pivot As IPoint _
            ) As IPoint
    If Pivot Is Nothing Then
        Align ComposedShapes, ActivePage, This.ReferencePoint
        Set AlignImposition = _
            FreePoint.Create( _
                ComposedShapes.LeftX, _
                ComposedShapes.TopY _
            )
    Else
        ComposedShapes.LeftX = Pivot.X
        ComposedShapes.TopY = Pivot.Y
    End If
End Function

Private Sub AddRandomElements( _
                ByVal ioElements As Collection, _
                ByVal Number As Long _
            )
    VBA.Randomize
    Dim Index As Long
    Dim RandomElement As IComposerElement
    Dim Shapes As ShapeRange
    Dim NewElement As IComposerElement
    For Index = 1 To Number
        Set RandomElement = This.Elements(RndInt(1, This.Elements.Count))
        Set Shapes = RandomElement.Shapes.Duplicate
        Set NewElement = ComposerElement.Create(Shapes)
        ioElements.Add NewElement
    Next Index
End Sub

Private Sub ProcessImpositions( _
                ByVal ImpositionRanges As Collection _
            )
    Dim ImpositionPages As Collection
    Set ImpositionPages = DistributeByMultiplePages(ImpositionRanges)
    If This.TwoSided Then Set ImpositionRanges = SeparateSides(ImpositionPages)
    If This.PlaceCropMarks Then PlaceCropMarksOnImpositions ImpositionRanges
End Sub

Private Sub PlaceCropMarksOnImpositions(ByVal ImpositionRanges As Collection)
    Dim Range As ShapeRange
    Dim Shape As Shape
    For Each Range In ImpositionRanges
        Range.FirstShape.Page.Activate
        With MarksSetter.Create(Range)
            .Bleeds = This.Bleeds
            .PlaceWhiteUndermark = False
            .Size = 3
            .SetMarksOnNewLayer
        End With
    Next Range
End Sub

Private Function DistributeByMultiplePages( _
                     ByVal ImpositionRanges As Collection _
                 ) As Collection
           
    Set DistributeByMultiplePages = New Collection
    Dim Index As Long
    Dim StartingPageIndex As Long
    StartingPageIndex = ActivePage.Index
    Dim Shapes As ShapeRange
    Dim Page As Page
    
    DistributeByMultiplePages.Add ActivePage
    If ImpositionRanges.Count > 1 Then _
        ActiveDocument.AddPages ImpositionRanges.Count - 1
    ActiveDocument.Pages(StartingPageIndex).Activate
    For Index = 2 To ImpositionRanges.Count
        Set Shapes = ImpositionRanges(Index)
        Set Page = ActiveDocument.Pages(Index + StartingPageIndex - 1)
        DistributeByMultiplePages.Add Page
        Shapes.MoveToLayer Page.ActiveLayer
    Next Index
            
End Function

Private Function SeparateSides( _
                     ByVal ImpositionPages As Collection _
                 ) As Collection
    Set SeparateSides = New Collection
    Dim Page As Page
    Dim TargetPage As Page
    For Each Page In ImpositionPages
        Set TargetPage = AddPage(Page)
        SeparateSide Page, TargetPage, SeparateSides
    Next Page
End Function

Private Sub SeparateSide( _
                ByVal SourcePage As Page, _
                ByVal TargetPage As Page, _
                ByVal ioSides As Collection _
            )
    Dim SideA As New ShapeRange
    SideA.AddRange SourcePage.Shapes.All
    Dim SideB As ShapeRange
    Set SideB = GetBSides(SourcePage.Shapes)
    SideA.RemoveRange SideB
    ioSides.Add SideA
    ioSides.Add SideB
    SourcePage.Activate
    Dim Shape As Shape
    For Each Shape In SideB
        Shape.RightX = _
            SourcePage.RightX - (Shape.LeftX - SourcePage.LeftX)
        Shape.MoveToLayer TargetPage.ActiveLayer
    Next Shape
End Sub

Private Property Get GetBSides(ByVal Shapes As Shapes) As ShapeRange
    Set GetBSides = Shapes.FindShapes(Name:=SIDE_B_MARK, Recursive:=False)
End Property

Private Function ShapesToElements(ByVal Shapes As ShapeRange) As Collection
    Dim Shape As Shape
    Set ShapesToElements = New Collection
    For Each Shape In Shapes
        ShapesToElements.Add ComposerElement.Create(Shape)
    Next Shape
End Function

Private Function ElementsToShapes( _
                     ByVal ComposerElements As Collection _
                 ) As ShapeRange
    Dim Item As IComposerElement
    Set ElementsToShapes = New ShapeRange
    For Each Item In ComposerElements
        ElementsToShapes.AddRange Item.Shapes
    Next Item
End Function

Private Function ImportAsMarkedElements(ByRef Motifs As Variant) As Collection
    Set ImportAsMarkedElements = New Collection
    Dim Item As Variant
    Dim Shapes As ShapeRange
    For Each Item In Motifs
        MotifOrThrow Item
        Set Shapes = ImportAsMarkedShapes(Item)
        AppendCollection _
            ImportAsMarkedElements, _
            DuplicateRangesAsElements(Shapes, Item.Quantity)
        RaiseEvent OnImportFile
    Next Item
End Function

Private Function ImportAsMarkedShapes( _
                     ByVal Motif As Motif _
                 ) As ShapeRange
    Dim ShapesToElement As ShapeRange
    Set ImportAsMarkedShapes = CreateShapeRange
    Dim ShapeA As Shape
    Dim ShapeB As Shape
    With Motif
        Set ShapeA = ImportSurface(.SurfaceA)
        ShapeA.Name = SIDE_A_MARK
        ImportAsMarkedShapes.Add ShapeA
        If This.TwoSided Then
            If .HasSurfaceB Then
                Set ShapeB = ImportSurface(.SurfaceB)
            Else
                Set ShapeB = ActiveLayer.CreateRectangleRect(ShapeA.BoundingBox)
                ShapeB.Fill.ApplyUniformFill CreateCMYKColor(0, 0, 0, 0)
            End If
            ShapeB.SetPosition ShapeA.PositionX, ShapeA.PositionY
            ShapeB.Name = SIDE_B_MARK
            ImportAsMarkedShapes.Add ShapeB
        End If
    End With
End Function

Private Function DuplicateRangesAsElements( _
                     ByVal Range As ShapeRange, _
                     ByVal Quantity As Long _
                 ) As Collection
    Set DuplicateRangesAsElements = New Collection
    DuplicateRangesAsElements.Add ComposerElement.Create(Range)
    If Quantity = 1 Then Exit Function
    Dim Index As Long
    For Index = 1 To Quantity - 1
        DuplicateRangesAsElements.Add ComposerElement.Create(Range.Duplicate)
    Next Index
End Function

Private Function ImportSurface(ByVal Surface As Surface) As Shape
    ActiveLayer.Import Surface.File
    Set ImportSurface = ActiveShape
End Function

Private Property Get GetMotifsCopy(ByRef Motifs As Variant) As Collection
    Set GetMotifsCopy = New Collection
    Dim Motif As Motif
    Dim Item As Variant
    For Each Item In Motifs
        MotifOrThrow Item
        Set Motif = Item.GetCopy
        GetMotifsCopy.Add Motif
    Next Item
End Property

Private Property Get BottomMarginY() As Double
    BottomMarginY = ActivePage.BottomY + This.EdgeMarginBottom
End Property
Private Property Get LeftMarginX() As Double
    LeftMarginX = ActivePage.LeftX + This.EdgeMarginLeft
End Property
Private Property Get RightMarginX() As Double
    RightMarginX = ActivePage.RightX - This.EdgeMarginRight
End Property
Private Property Get TopMarginY() As Double
    TopMarginY = ActivePage.TopY - This.EdgeMarginTop
End Property

Public Sub MotifOrThrow(ByRef MaybeMotif As Variant)
    If Not AssignedObject(MaybeMotif) Then GoTo Fail
    If Not TypeOf MaybeMotif Is Motif Then GoTo Fail
    Exit Sub
Fail:
    Throw "Motifs ?????? ????????? ??????? ???? Motif"
End Sub

