VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sticker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===============================================================================
' # Declarations

Private Type typeThis
    CutColor As Color
    Distance As Double
    MarksColor As Color
    PerfcutColor As Color
    SeparationsShift As Double
    ShapesToCut As ShapeRange
    ShapesToPrint As ShapeRange
    TargetCutColor As Color
    TargetPerfcutColor As Color
    WhiteColor As Color
End Type
Private This As typeThis

Private Const MARK_LESSER_SIZE = 3#
Private Const BAR_EDGE_OFFSET = 10#
Private Const WHITE_MARKS_BLEED = 1#

'===============================================================================
' # Handlers

Private Sub Class_Initialize()
    With This
        Set .CutColor = CreateColor("SPOT,USER,2,100,ee33f031-17c3-40a4-8e1a-e2fc4267738a,~,CutContour,1,LAB,USER,138,74,-8,100,00000000-0000-0000-0000-000000000000,~")
        Set .PerfcutColor = CreateColor("SPOT,USER,0,100,ee33f031-17c3-40a4-8e1a-e2fc4267738a,~,CutContour Perfcut,1,LAB,USER,110,15,-40,100,00000000-0000-0000-0000-000000000000,~")
        Set .WhiteColor = CreateColor("SPOT,USER,1,100,ee33f031-17c3-40a4-8e1a-e2fc4267738a,~,White,1,LAB,USER,231,-4,75,100,00000000-0000-0000-0000-000000000000,~")
        Set .TargetCutColor = CreateCMYKColor(0, 0, 0, 100)
        Set .TargetPerfcutColor = CreateRGBColor(255, 0, 0)
        Set .MarksColor = CreateCMYKColor(0, 0, 0, 100)
        .SeparationsShift = 10
    End With
End Sub

Private Sub Class_Terminate()
'
End Sub

'===============================================================================
' # Public

Public Property Get GetShapesToCut() As ShapeRange
    Set GetShapesToCut = This.ShapesToCut
End Property

Public Property Get GetShapesToPrint() As ShapeRange
    Set GetShapesToPrint = This.ShapesToPrint
End Property

Public Sub Separate(ByVal Shapes As ShapeRange)

    Set Shapes = Shapes.UngroupAllEx
    
    Dim Box As Rect
    Set Box = Shapes.BoundingBox
    With This
        Box.Inflate .Distance, 0, .Distance, .Distance
    End With
    
    Dim BlackMarks As Shape
    Set BlackMarks = DrawBlackMarks(Box)
    Dim WhiteMarks As Shape
    Set WhiteMarks = DrawWhiteMarks(Box)
    WhiteMarks.OrderBackOf BlackMarks
    
    Dim AllShapes As New ShapeRange
    AllShapes.AddRange Shapes
    AllShapes.Add BlackMarks
    AllShapes.Add WhiteMarks
    
    Dim CutShapes As ShapeRange
    Set CutShapes = FindCutShapes(AllShapes)
    'Debug.Print "CutShapes = " & CutShapes.Count
    Dim PerfcutShapes As ShapeRange
    Set PerfcutShapes = FindPerfcutShapes(AllShapes)
    'Debug.Print "PerfcutShapes = " & PerfcutShapes.Count
    Dim WhiteShapes As ShapeRange
    Set WhiteShapes = FindWhiteShapes(AllShapes)
    'Debug.Print "WhiteShapes = " & WhiteShapes.Count
    SetOverprint WhiteShapes
    
    Dim ToPrint As New ShapeRange
    ToPrint.AddRange AllShapes
    ToPrint.RemoveRange CutShapes
    ToPrint.RemoveRange PerfcutShapes
    
    Dim ToCut As New ShapeRange
    ToCut.AddRange CutShapes
    ToCut.AddRange PerfcutShapes
    ToCut.Add BlackMarks
    FindCutShapes(ToCut).SetOutlineProperties Color:=This.TargetCutColor
    FindPerfcutShapes(ToCut).SetOutlineProperties Color:=This.TargetPerfcutColor
        
    With This
        Set .ShapesToPrint = ToPrint.Duplicate
        Set .ShapesToCut = ToCut.Duplicate
        .ShapesToCut.Move _
            0, _
            -( _
                AverageDim(.ShapesToCut) * (.SeparationsShift / 100) _
                + .ShapesToCut.SizeHeight _
            )
        AllShapes.Delete
    End With

End Sub

Public Sub SetDistance(ByVal Distance As Double)
    This.Distance = Distance
End Sub

Public Sub SetMainColors( _
               ByVal CutColor As Color, _
               ByVal PerfcutColor As Color, _
               ByVal WhiteColor As Color _
            )
    With This
        Set .CutColor = CreateColor
        .CutColor.CopyAssign CutColor
        Set .PerfcutColor = CreateColor
        .PerfcutColor.CopyAssign PerfcutColor
        Set .WhiteColor = CreateColor
        .WhiteColor.CopyAssign WhiteColor
    End With
End Sub

Public Sub SetSeparationsShift(ByVal SeparationsShift As Double)
    This.SeparationsShift = SeparationsShift
End Sub

Public Property Get Self() As Sticker
    Set Self = Me
End Property

'===============================================================================
' # Helpers

Private Function DrawBlackMarks(ByVal Box As Rect) As Shape
    Dim Result As New ShapeRange
    
    Dim MarksCol As New ShapeRange
    MarksCol.Add DrawSquare(Box.Left - MARK_LESSER_SIZE, Box.Top)
    MarksCol.Add _
        DrawSquare( _
            Box.Left - MARK_LESSER_SIZE, _
            Box.Bottom - MARK_LESSER_SIZE _
        )
    MarksCol.Add _
        DrawSquare( _
            Box.Left - MARK_LESSER_SIZE, _
            MarksCol.CenterY - MARK_LESSER_SIZE / 2 _
        )
    
    Dim MarksColRight As ShapeRange
    Set MarksColRight = MarksCol.Duplicate
    MarksColRight.LeftX = Box.Right
    
    Result.AddRange MarksCol
    Result.AddRange MarksColRight
    
    Dim Bar As Shape
    Set Bar = _
        ActiveLayer.CreateRectangle( _
            Box.Left + BAR_EDGE_OFFSET, _
            Box.Bottom, _
            Box.Right - BAR_EDGE_OFFSET, _
            Box.Bottom - MARK_LESSER_SIZE _
        )
    Bar.Fill.ApplyUniformFill This.MarksColor
    Bar.Outline.SetNoOutline
    Result.Add Bar
    Set DrawBlackMarks = Result.Combine
End Function

Private Function DrawWhiteMarks(ByVal Box As Rect) As Shape
    Dim Result As New ShapeRange
    
    Dim Vertical As Shape
    Set Vertical = _
        ActiveLayer.CreateRectangle( _
            Box.Left - MARK_LESSER_SIZE - WHITE_MARKS_BLEED, _
            Box.Top + MARK_LESSER_SIZE + WHITE_MARKS_BLEED, _
            Box.Left + WHITE_MARKS_BLEED, _
            Box.Bottom - MARK_LESSER_SIZE - WHITE_MARKS_BLEED _
        )
    Result.Add Vertical
    Dim WhiteWidth As Double
    WhiteWidth = Vertical.SizeWidth
    Result.Add Vertical.Duplicate(Box.Width + WhiteWidth - WHITE_MARKS_BLEED * 2, 0)
    
    Dim Horizontal As Shape
    Set Horizontal = _
        ActiveLayer.CreateRectangle( _
            Result.LeftX, _
            Box.Bottom + WHITE_MARKS_BLEED, _
            Result.RightX, _
            Box.Bottom - MARK_LESSER_SIZE - WHITE_MARKS_BLEED _
        )
    Result.Add Horizontal
    
    Set DrawWhiteMarks = Weld(Result)
    DrawWhiteMarks.Fill.ApplyUniformFill This.WhiteColor
    DrawWhiteMarks.Outline.SetNoOutline
End Function

Private Function DrawSquare( _
                     ByVal LeftX As Double, _
                     ByVal BottomY As Double _
                 ) As Shape
    With This
        Set DrawSquare = _
            ActiveLayer.CreateRectangle( _
                LeftX, _
                BottomY + MARK_LESSER_SIZE, _
                LeftX + MARK_LESSER_SIZE, _
                BottomY _
            )
        DrawSquare.Fill.ApplyUniformFill .MarksColor
        DrawSquare.Outline.SetNoOutline
    End With
End Function

Private Property Get FindCutShapes(ByVal Shapes As ShapeRange) As ShapeRange
    Set FindCutShapes = FindShapesByOutlineColor(Shapes, This.CutColor)
End Property

Private Property Get FindPerfcutShapes(ByVal Shapes As ShapeRange) As ShapeRange
    Set FindPerfcutShapes = FindShapesByOutlineColor(Shapes, This.PerfcutColor)
End Property

Private Property Get FindWhiteShapes(ByVal Shapes As ShapeRange) As ShapeRange
    Set FindWhiteShapes = FindShapesByFillColor(Shapes, This.WhiteColor)
End Property

Private Sub SetOverprint(ByVal Shapes As ShapeRange)
    Dim Shape As Shape
    For Each Shape In Shapes
        Shape.OverprintFill = True
    Next Shape
End Sub
