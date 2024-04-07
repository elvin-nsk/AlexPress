VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExportPdfView 
   Caption         =   "Постраничный экспорт PDF"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3750
   OleObjectBlob   =   "ExportPdfView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExportPdfView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================

Public IsExportToPageSize As Boolean
Public IsCancel As Boolean

Private Type This
    PdfPath As FolderBrowserHandler
End Type
Private This As This

'===============================================================================

Private Sub UserForm_Initialize()
    Set This.PdfPath = _
        FolderBrowserHandler.New_(TextBoxPdfPath, ButtonBrowsePdfPath)
End Sub

Private Sub UserForm_Activate()
    '
End Sub

Private Sub ExportToContentSize_Click()
    IsExportToPageSize = False
    Hide
End Sub

Private Sub ExportToPageSize_Click()
    IsExportToPageSize = True
    Hide
End Sub

'===============================================================================

Private Sub FormCancel()
    Hide
    IsCancel = True
End Sub

'===============================================================================

Private Sub UserForm_QueryClose(Сancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Сancel = True
        FormCancel
    End If
End Sub
