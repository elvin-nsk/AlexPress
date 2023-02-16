VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainView 
   Caption         =   "AlexPress"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7785
   OleObjectBlob   =   "MainView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================

Public BleedsMin As Double
Public BleedsMax As Double

Private Type typeThis
    Main As MainLogic
    Motifs As MotifsCollector
    ActiveMotif As Motif
    QuantityOnPressSheet As TextBoxHandler
    PressSheetWidth As TextBoxHandler
    PressSheetHeight As TextBoxHandler
    ImpositionColumns As TextBoxHandler
    ImpositionRows As TextBoxHandler
    ImpositionCreated As Boolean
End Type
Private This As typeThis

Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1

'===============================================================================

Private Sub UserForm_Initialize()
    Set App = Application
    Set This.Motifs = New MotifsCollector
    Set This.Main = MainLogic.Create(Me)
    With This
        Set This.QuantityOnPressSheet = _
            TextBoxHandler.Create(QuantityOnPressSheet, TextBoxTypeLong)
        Set This.PressSheetWidth = _
            TextBoxHandler.Create(PressSheetWidth, TextBoxTypeLong)
        Set This.PressSheetHeight = _
            TextBoxHandler.Create(PressSheetHeight, TextBoxTypeLong)
        Set This.ImpositionColumns = _
            TextBoxHandler.Create(ImpositionColumns, TextBoxTypeLong)
        Set This.ImpositionRows = _
            TextBoxHandler.Create(ImpositionRows, TextBoxTypeLong)
    End With
End Sub

Private Sub UserForm_Activate()
    This.QuantityOnPressSheet = 1
    ResetStatus
    CheckControls
End Sub

Private Sub tbBleeds_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    OnlyNum KeyAscii
End Sub
Private Sub tbBleeds_AfterUpdate()
    CheckRange tbBleeds, BleedsMin, BleedsMax
End Sub

Private Sub tbTrim_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    OnlyInt KeyAscii
End Sub
Private Sub tbTrim_AfterUpdate()
    CheckRange tbTrim, 1, 10
End Sub

Private Sub tbPages_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    OnlyNum KeyAscii
End Sub
Private Sub tbPages_AfterUpdate()
    CheckRange tbPages, 1
End Sub

Private Sub btnCancel_Click()
    FormCancel
End Sub

Private Sub btnAddBleeds_Click()
    This.Main.AddBleeds Me
End Sub

Private Sub btnAddPages_Click()
    This.Main.AddPages Me
End Sub

Private Sub SetFace_Click()
    CreateMotifIfNothing
    With This.ActiveMotif
        Set .SurfaceA = New Surface
        Set .SurfaceA.Content = ActiveSelectionRange.FirstShape
        .Quantity = This.QuantityOnPressSheet
    End With
    If OptionNoBacks Then
        ActiveMotifAddAndReset
    Else
        CancelAssignment.Enabled = True
        If This.ActiveMotif.HasSurfaceB Then
            ActiveMotifAddAndReset
        Else
            SetMotifStateToNeedBack
        End If
    End If
End Sub

Private Sub SetBack_Click()
    CreateMotifIfNothing
    With This.ActiveMotif
        Set .SurfaceB = New Surface
        Set .SurfaceB.Content = ActiveSelectionRange.FirstShape
        .Quantity = This.QuantityOnPressSheet
    End With
    If This.ActiveMotif.HasSurfaceA Then
        ActiveMotifAddAndReset
    Else
        CancelAssignment.Enabled = True
        SetMotifStateToNeedFace
    End If
End Sub

Private Sub CancelAssignment_Click()
    If Not This.ActiveMotif Is Nothing Then
        Set This.ActiveMotif = Nothing
        SetFace.Enabled = True
        SetBack.Enabled = True
        MotifStatusLabel.Caption = vbNullString
    End If
End Sub

Private Sub ResetAssignments_Click()
    ResetAll
End Sub

Private Sub CreateImpositions_Click()
    This.Main.MakeImposition Me, This.Motifs
    This.ImpositionCreated = True
End Sub

Private Sub SwapPressSheetDims_Click()
    Dim Temp As Long
    Temp = PressSheetWidth
    PressSheetWidth = PressSheetHeight
    PressSheetHeight = Temp
End Sub

Private Sub SwapRowsAndColumns_Click()
    Dim Temp As Long
    Temp = ImpositionColumns
    ImpositionColumns = ImpositionRows
    ImpositionRows = Temp
End Sub

Private Sub App_SelectionChange()
    CheckControls
End Sub

'===============================================================================

Private Sub CheckControls()
    ExecutabilityCheck
    VisibilityCheck
    ImpositionCheck
End Sub

Private Sub ExecutabilityCheck()
    If ActiveDocument Is Nothing Then
        btnAddBleeds.Enabled = False
        btnAddBleeds.Caption = "Нет документа"
        btnAddPages.Enabled = False
        btnAddPages.Caption = "Нет документа"
        Exit Sub
    End If
    btnAddPages.Enabled = True
    btnAddPages.Caption = "Добавить страницы"
    If ActiveSelectionRange.Count > 1 Then
        btnAddBleeds.Enabled = False
        btnAddBleeds.Caption = "Несколько объектов"
        Exit Sub
    End If
    If ActiveSelectionRange.Count < 1 Then
        btnAddBleeds.Enabled = False
        btnAddBleeds.Caption = "Не выбран объект"
        Exit Sub
    End If
    btnAddBleeds.Enabled = True
    btnAddBleeds.Caption = "Добавить припуски"
End Sub

Private Sub VisibilityCheck()
    If ActiveDocument Is Nothing Then
        RastrHide
        Exit Sub
    End If
    If IsShapeType(ActiveSelectionRange.FirstShape, cdrBitmapShape) Then
        RastrShow
    Else
        RastrHide
    End If
End Sub

Private Sub RastrShow()
    cbTrim.Enabled = True
    tbTrim.Enabled = True
    lblPix.Enabled = True
    cbFlatten.Enabled = True
End Sub

Private Sub RastrHide()
    cbTrim.Enabled = False
    tbTrim.Enabled = False
    lblPix.Enabled = False
    cbFlatten.Enabled = False
End Sub

Private Sub ImpositionCheck()
    If ActiveDocument Is Nothing Then
        DisableMotifControls
        ResetStatus
        Exit Sub
    End If
    UpdateImpositionState
    If ActiveSelectionRange.Count = 0 Then
        MotifStatusLabel.Caption = vbNullString
        DisableMotifControls
        Exit Sub
    End If
    If ActiveSelectionRange.Count > 1 Then
        SetMotifStatusSelectOne
        DisableMotifControls
        Exit Sub
    End If
    If This.ActiveMotif Is Nothing Then
        SetFace.Enabled = True
        SetBack.Enabled = OptionWithBacks
        CancelAssignment.Enabled = False
        MotifStatusLabel.Caption = vbNullString
    Else
        If OptionWithBacks Then
            If This.ActiveMotif.HasSurfaceA _
           And Not This.ActiveMotif.HasSurfaceB Then
                CancelAssignment.Enabled = True
                SetMotifStateToNeedBack
            ElseIf Not This.ActiveMotif.HasSurfaceA _
               And This.ActiveMotif.HasSurfaceB Then
                CancelAssignment.Enabled = True
                SetMotifStateToNeedFace
            Else
                MotifStatusLabel.Caption = vbNullString
            End If
        Else
            MotifStatusLabel.Caption = vbNullString
        End If
    End If
    
    

End Sub

Private Sub CreateMotifIfNothing()
    If This.ActiveMotif Is Nothing Then Set This.ActiveMotif = New Motif
End Sub

Private Property Get MotifAddingHasBegun() As Boolean
    MotifAddingHasBegun = (This.Motifs.Count > 0) _
                       Or (Not This.ActiveMotif Is Nothing)
End Property

Private Property Get UnsafeToClose() As Boolean
    UnsafeToClose = MotifAddingHasBegun And Not This.ImpositionCreated
End Property

Private Sub DisableMotifControls()
    SetFace.Enabled = False
    SetBack.Enabled = False
    CancelAssignment.Enabled = False
End Sub

Private Sub ActiveMotifAddAndReset()
    This.Motifs.Add This.ActiveMotif
    Set This.ActiveMotif = Nothing
    SetMotifStateToAdded
    UpdateImpositionState
End Sub

Private Sub ResetAll()
    Set This.Motifs = New MotifsCollector
    UpdateImpositionState
End Sub

Private Sub SetMotifStateToNeedFace()
    SetFace.Enabled = True
    SetBack.Enabled = False
    MotifStatusLabel.Caption = "Оборот добавлен, добавьте лицо"
End Sub

Private Sub SetMotifStateToNeedBack()
    SetFace.Enabled = False
    SetBack.Enabled = True
    MotifStatusLabel.Caption = "Лицо добавлено, добавьте оборот"
End Sub

Private Sub SetMotifStateToAdded()
    DisableMotifControls
    MotifStatusLabel.Caption = "Макет добавлен"
End Sub

Private Sub SetMotifStatusAreadyAdded()
    MotifStatusLabel.Caption = "Этот макет уже добавлен"
End Sub

Private Sub SetMotifStatusSelectOne()
    MotifStatusLabel.Caption = "Выберите один объект"
End Sub

Private Sub UpdateImpositionState()
    If This.Motifs.Count > 0 Then
        ImpositionStatusLabel.Caption = "Кол-во макетов: " & This.Motifs.Count
        ResetAssignments.Enabled = True
        CreateImpositions.Enabled = True
    Else
        ResetStatus
        ResetAssignments.Enabled = False
        CreateImpositions.Enabled = False
    End If
End Sub

Private Sub ResetStatus()
    MotifStatusLabel.Caption = vbNullString
    ImpositionStatusLabel.Caption = vbNullString
End Sub

Private Sub FormCancel()
    Dim OkToExit As VbMsgBoxResult
    If UnsafeToClose Then
        OkToExit = AskToClose
    Else
        OkToExit = vbOK
    End If
    If OkToExit = vbOK Then
        This.Main.Dispose Me
        Me.Hide
    End If
End Sub

Private Function AskToClose() As VbMsgBoxResult
    AskToClose = _
        VBA.MsgBox( _
            "Макеты не разложены." _
          & vbCr _
          & "Если закрыть, назначение макетов сбросится." _
          & vbCr _
          & "Хотите закрыть?", _
            vbOKCancel + vbExclamation _
        )
End Function

'===============================================================================

Private Sub OnlyInt(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub OnlyNum(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Asc(",")
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub CheckRange(TextBox As MSForms.TextBox, ByVal Min As Double, Optional ByVal Max As Double = 2147483647)
    With TextBox
        If CDbl(.Value) > Max Then .Value = CStr(Max)
        If CDbl(.Value) < Min Then .Value = CStr(Min)
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        FormCancel
    End If
End Sub
