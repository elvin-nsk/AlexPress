VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MotifsCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
' # Declarations

Private Type typeThis
    Motifs As Collection
End Type
Private This As typeThis

'===============================================================================
' # Handlers

Private Sub Class_Initialize()
    Set This.Motifs = New Collection
End Sub

'===============================================================================
' # Public

Public Sub Add(ByVal Motif As Motif)
    This.Motifs.Add Motif
End Sub

Public Property Get AsCollection() As Collection
    Set AsCollection = This.Motifs
End Property

Public Property Get Count() As Long
    Count = This.Motifs.Count
End Property

Public Function ExportToTemp() As Collection
    Dim TempFolder As String
    TempFolder = GetTempFolder
    If TempFolder = vbNullString Then Exit Function
    
    Set ExportToTemp = New Collection
    
    Dim Motif As Motif
    For Each Motif In This.Motifs
        With Motif
            If .HasSurfaceA Then _
                ExportSurface TempFolder, Motif.SurfaceA, ExportToTemp
            If .HasSurfaceB Then _
                ExportSurface TempFolder, Motif.SurfaceB, ExportToTemp
        End With
    Next Motif
End Function

Private Function ExportSurface( _
                     ByVal TempFolder As String, _
                     ByVal ioSurface As Surface, _
                     ByVal ioCollectionOfFiles As Collection _
                 )
    Dim File As IFileSpec
    Set File = FileSpec.New_
    File.Path = TempFolder
    File.Ext = "cdr"
    File.NameWithoutExt = CreateGUID
    Dim Options As StructSaveAsOptions
    Set Options = CreateStructSaveAsOptions
    Options.Range = cdrSelection
    With ioSurface
        .Content.Page.Parent.Parent.Activate
        .Content.CreateSelection
        ActiveDocument.SaveAsCopy File, Options
        Set .File = File
        ioCollectionOfFiles.Add File
    End With
End Function

Public Property Get FindIndexByShape(ByVal Shape As Shape) As Long
    With This
        If .Motifs.Count = 0 Then Exit Property
        Dim Index As Long
        For Index = 1 To This.Motifs.Count
            If MotifContainsShape(.Motifs(Index), Shape) Then
                FindIndexByShape = Index
                Exit Property
            End If
        Next Index
    End With
End Property

Public Property Get Item(ByVal Index As Long) As Motif
    Set Item = This.Motifs(Index)
End Property

Public Sub Pop()
    Remove Count
End Sub

Public Function Remove(ByVal Index As Long)
    This.Motifs.Remove Index
End Function

Public Property Get Self() As MotifsCollector
    Set Self = Me
End Property

'===============================================================================
' # Helpers

Private Property Get MotifContainsShape( _
                         ByVal Motif As Motif, _
                         ByVal Shape As Shape _
                     ) As Boolean
    If Motif Is Nothing Then Exit Property
    If SurfaceContainsShape(Motif.SurfaceA, Shape) Then
        MotifContainsShape = True
        Exit Property
    End If
    If SurfaceContainsShape(Motif.SurfaceB, Shape) Then _
        MotifContainsShape = True
End Property

Private Property Get SurfaceContainsShape( _
                         ByVal Surface As Surface, _
                         ByVal Shape As Shape _
                     ) As Boolean
    If Surface Is Nothing Then Exit Property
    SurfaceContainsShape = Surface.Content Is Shape
End Property
