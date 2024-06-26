VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' Config v2022.12.22
' ��������, �������� � ���������� ��������
'===============================================================================

'@PredeclaredId
Option Explicit

#Const ReadOnly = False
#Const DefaultContent = False

'===============================================================================

Private Const iniFileName As String = "elvin_" & APP_NAME & ".ini"

'===============================================================================

Public Bleeds As Double
Public BitmapTrim As Boolean
Public BitmapTrimSize As Long
Public BitmapFlatten As Boolean
Public RoundSize As Boolean
Public RoundDec As Long

Public PagesToAdd As Long

Public MotifsSpaces As Double
Public OptionCutmarks As Boolean
Public OptionNoBacks As Boolean
Public OptionNoMarks As Boolean
Public OptionPlotter As Boolean
Public OptionRastr As Boolean
Public OptionWithBacks As Boolean
Public PressSheetHeight As Double
Public PressSheetSpaces As Double
Public PressSheetWidth As Double

Public StickerSpace As Double

Public WorkLogsPath As String
Public WorkLogsEnabled As Boolean

Public ExportPdfPath As String

'-------------------------------------------------------------------------------

Private Sub Assign(ByVal Action As iniAction)

    With iniFile.Create
    
        iniFilePrepare .Self, Action
        
        .Section = "Bleeds"
        .KeyVar "Bleeds", Bleeds, 3
        .KeyVar "BitmapTrim", BitmapTrim, True
        .KeyVar "BitmapTrimSize", BitmapTrimSize, 2
        .KeyVar "BitmapFlatten", BitmapFlatten, False
        .KeyVar "RoundSize", RoundSize, True
        .KeyVar "RoundDec", RoundDec, 0
        
        .Section = "Pages"
        .KeyVar "PagesToAdd", PagesToAdd, 1
        
        .Section = "Imposition"
        .KeyVar "MotifsSpaces", MotifsSpaces, 0
        .KeyVar "OptionCutmarks", OptionCutmarks, True
        .KeyVar "OptionNoBacks", OptionNoBacks, False
        .KeyVar "OptionNoMarks", OptionNoMarks, False
        .KeyVar "OptionPlotter", OptionPlotter, False
        .KeyVar "OptionRastr", OptionRastr, False
        .KeyVar "OptionWithBacks", OptionWithBacks, True
        .KeyVar "PressSheetHeight", PressSheetHeight, 450
        .KeyVar "PressSheetSpaces", PressSheetSpaces, 0
        .KeyVar "PressSheetWidth", PressSheetWidth, 320
        
        .Section = "Sticker"
        .KeyVar "StickerSpace", StickerSpace, 3
        
        .Section = "WorkLogger"
        .KeyVar "WorkLogsPath", WorkLogsPath, "c:\temp\"
        .KeyVar "WorkLogsEnabled", WorkLogsEnabled, False
        
        .Section = "ExportPdf"
        .KeyVar "ExportPdfPath", ExportPdfPath, "c:\temp\"
    
    End With
    
End Sub

'-------------------------------------------------------------------------------

'�����������. ������ ini-���� ��������� ���������� (����. � ����������).
#If DefaultContent Then
Private Sub MakeDefaultFile(ByVal File As String)

    Dim Content As String
    
    Content = Content & "; ��������� Windows-1251" & vbCrLf
    Content = Content & "" & vbCrLf
    Content = Content & "[DEFAULT]" & vbCrLf
    Content = Content & "" & vbCrLf
    Content = Content & "; �����" & vbCrLf
    Content = Content & "Bulka=������" & vbCrLf
    Content = Content & "" & vbCrLf
    Content = Content & "; �����" & vbCrLf
    Content = Content & "Palka=1" & vbCrLf
    Content = Content & "" & vbCrLf
    Content = Content & "; ������" & vbCrLf
    Content = Content & "Kryshka=2" & vbCrLf
    
    WriteFile Content, File
    
End Sub
#End If

'===============================================================================
'           * * * �������, �� ��������������� ��� ��������� * * *
'===============================================================================
' # Constructors

Public Function Bind_() As Config
    Set Bind_ = New Config
    Bind_.ForceLoad
End Function

#If Not ReadOnly Then
Private Sub Class_Terminate()
    ForceSave
End Sub
#End If

'===============================================================================
' # Public

Public Sub ForceLoad()
    Assign iniRead
End Sub

#If Not ReadOnly Then
Public Sub ForceSave()
    Assign iniWrite
End Sub
#End If

Public Property Get Self() As Config
    Set Self = Me
End Property

'===============================================================================

Private Sub iniFilePrepare(ByVal iniFile As iniFile, ByVal Action As iniAction)

    Dim iniFilePath As String

    With iniFile
        Select Case True
        Case FileExists(GMSManager.UserGMSPath & iniFileName)
            .File = Application.GMSManager.UserGMSPath & iniFileName
            .Action = Action
        Case FileExists(GMSManager.GMSPath & iniFileName)
            .File = Application.GMSManager.GMSPath & iniFileName
            .Action = Action
        Case Else
            .File = GMSManager.UserGMSPath & iniFileName
            #If DefaultContent Then
            MakeDefaultFile .File
            If Action = iniRead Then
                .Action = iniRead
            Else
                .Action = iniWrite
            End If
            #Else
            If Action = iniRead Then .Action = iniReset Else .Action = iniWrite
            #End If
        End Select
    End With

End Sub

Private Function FileExists(ByVal File As String) As Boolean
    If File = "" Then Exit Function
    FileExists = VBA.Len(VBA.Dir(File)) > 0
End Function

Private Sub WriteFile( _
               ByVal Content As String, _
               ByVal File As String _
           )
    Dim FileNum As Long
    FileNum = VBA.FreeFile
    Open File For Output Access Write As #FileNum
    Print #FileNum, Content
    Close #FileNum
End Sub
