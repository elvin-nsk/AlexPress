VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettingsView 
   Caption         =   "���������"
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
    WorkLogsPath As FolderBrowserHandler
End Type
Private This As typeThis

'===============================================================================

Private Sub UserForm_Initialize()
    Set This.WorkLogsPath = _
        FolderBrowserHandler.New_(TextBoxWorkLogsPath, ButtonBrowseWorkLogsPath)
End Sub

Private Sub UserForm_Activate()
    '
End Sub

Private Sub btnClose_Click()
    FormClose
End Sub

'===============================================================================

Private Sub FormClose()
    Me.Hide
End Sub

'===============================================================================

Private Sub UserForm_QueryClose(�ancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        �ancel = True
        FormClose
    End If
End Sub
