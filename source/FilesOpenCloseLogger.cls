VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilesOpenCloseLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

'===============================================================================
' # Declarations

Private Type typeThis
    Path As String
    DocsTimes As IList
End Type
Private This As typeThis

Private WithEvents HostEvents As Application
Attribute HostEvents.VB_VarHelpID = -1

'===============================================================================
' # Constructors

Public Function New_( _
                    ByVal Path As String _
                ) As FilesOpenCloseLogger
    Set New_ = New FilesOpenCloseLogger
    New_.Inject Path
End Function

Friend Sub Inject( _
               ByVal Path As String _
           )
    With This
        .Path = Path
        Set .DocsTimes = List.New_
        Set HostEvents = Application
    End With
End Sub

'===============================================================================
' # Event handlers

Private Sub HostEvents_DocumentClose(ByVal Doc As Document)
    On Error GoTo Catch

    Dim Time As Variant
    Time = GetDocTimeAndRemoveFromList(Doc)
    If VBA.IsEmpty(Time) Then
        Debug.Print "Warning: Time пустое для файла " & Doc.Name
        Exit Sub
    End If
    WriteLogFile Doc.FileName, VBA.DateDiff("n", Time, VBA.Now)
    
Finally:
    Exit Sub

Catch:
    Debug.Print "Error: " & VBA.Err.Description
    Resume Finally
End Sub

Private Sub HostEvents_DocumentOpen( _
                ByVal Doc As Document, _
                ByVal FileName As String _
            )
    On Error GoTo Catch
            
    Dim Item As New DocTime
    Set Item.Doc = Doc
    Item.Time = VBA.Now
    This.DocsTimes.Add Item
    
Finally:
    Exit Sub

Catch:
    Debug.Print "Error: " & VBA.Err.Description
    Resume Finally
End Sub

Private Sub Class_Terminate()
    Dispose
End Sub

'===============================================================================
' # Public

Public Sub Dispose()
    Set HostEvents = Nothing
End Sub

Public Property Get Self() As FilesOpenCloseLogger
    Set Self = Me
End Property

'===============================================================================
' # Helpers

Private Function GetDocTimeAndRemoveFromList( _
                     ByVal Doc As Document _
                 ) As Variant
    Dim Item As DocTime
    For Each Item In This.DocsTimes.Items
        If Item.Doc Is Doc Then
            GetDocTimeAndRemoveFromList = Item.Time
            This.DocsTimes.Remove This.DocsTimes.IndexOf(Item)
            Exit Function
        End If
    Next Item
End Function

Private Sub WriteLogFile(ByVal FileName As String, ByVal Minutes As Long)
    Dim File As IFileSpec
    Set File = FileSpec.New_
    File.Path = This.Path
    File.Name = FileName
    File.Ext = "txt"
    CheckForExistanceAndRename File
    WriteFile _
        Minutes & " мин.", _
        File
End Sub

Private Function CheckForExistanceAndRename(ByVal ioFile As IFileSpec)
    Dim Counter As Long
    Do While FileExists(ioFile) And Counter < 1000
        ioFile.NameWithoutExt = ioFile.NameWithoutExt & "+"
        Counter = Counter + 1
    Loop
    'Debug.Print Counter
End Function

Private Property Get ToHoursAndMinutes(ByVal Minutes As Long) As String
  ToHoursAndMinutes = _
     VBA.Fix(Minutes / 60) & " час. " & Minutes Mod 60 & " мин."
End Property
