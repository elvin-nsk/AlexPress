VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NumberSetterView 
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4695
   OleObjectBlob   =   "NumberSetterView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NumberSetterView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================

Public IsOk As Boolean
Public IsCancel As Boolean
Public Prefix As TextBoxHandler
Public StartingNumber As TextBoxHandler

'===============================================================================

Private Sub UserForm_Initialize()
    Caption = "NumberSetter"
    tbStartingNumber = 1
    Set Prefix = _
        TextBoxHandler.Create(tbPrefix, TextBoxTypeString)
    Set StartingNumber = _
        TextBoxHandler.Create(tbStartingNumber, TextBoxTypeLong, 1)
End Sub

Private Sub UserForm_Activate()
    '
End Sub

Private Sub btnOk_Click()
    FormŒ 
End Sub

Private Sub btnCancel_Click()
    FormCancel
End Sub

'===============================================================================

Private Sub FormŒ ()
    Me.Hide
    IsOk = True
End Sub

Private Sub FormCancel()
    Me.Hide
    IsCancel = True
End Sub

'===============================================================================

Private Sub UserForm_QueryClose(—ancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        —ancel = True
        FormCancel
    End If
End Sub
