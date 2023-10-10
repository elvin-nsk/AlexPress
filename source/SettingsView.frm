VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettingsView 
   Caption         =   "Íàñòðîéêè"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3750
   OleObjectBlob   =   "SettingsView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SettingsView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================

Private Type typeThis
    WorkLogsPath As TextBoxHandler
End Type
Private This As typeThis

'===============================================================================

Private Sub UserForm_Initialize()
    Set This.WorkLogsPath = _
        TextBoxHandler.New_(TextBoxWorkLogsPath, TextBoxTypeString)
End Sub

Private Sub UserForm_Activate()
    '
End Sub

Private Sub ButtonBrowseWorkLogsPath_Click()
    Dim LastPath As String
    LastPath = This.WorkLogsPath
    Dim Folder As IFileSpec
    Set Folder = FileSpec.New_(This.WorkLogsPath)
    Folder.Path = CorelScriptTools.GetFolder(Folder.Path)
    If Folder.Path = "\" Then
        This.WorkLogsPath = LastPath
    Else
        This.WorkLogsPath = Folder.Path
    End If
End Sub

Private Sub btnClose_Click()
    FormClose
End Sub

'===============================================================================

Private Sub FormClose()
    Me.Hide
End Sub

'===============================================================================

Private Sub UserForm_QueryClose(Ñancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Ñancel = True
        FormClose
    End If
End Sub
