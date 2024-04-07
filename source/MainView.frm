VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainView 
   Caption         =   "AlexPress"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7755
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
    ImpositionCreated As Boolean
    MotifsSpaces As TextBoxHandler
    PressSheetWidth As TextBoxHandler
    PressSheetHeight As TextBoxHandler
    PressSheetSpaces As TextBoxHandler
    StickerSpace As TextBoxHandler
    QuantityOnPressSheet As TextBoxHandler
End Type
Private This As typeThis

Private WithEvents Host As Application
Attribute Host.VB_VarHelpID = -1

'===============================================================================
' # Event handlers

Private Sub UserForm_Initialize()
    Set Host = Application
    Set This.Motifs = New MotifsCollector
    Set This.Main = MainLogic.New_(Me)
    With This
        Set This.MotifsSpaces = _
            TextBoxHandler.New_(MotifsSpaces, TextBoxTypeDouble, 0)
        Set This.PressSheetWidth = _
            TextBoxHandler.New_(PressSheetWidth, TextBoxTypeDouble, 1)
        Set This.PressSheetHeight = _
            TextBoxHandler.New_(PressSheetHeight, TextBoxTypeDouble, 1)
        Set This.PressSheetSpaces = _
            TextBoxHandler.New_(PressSheetSpaces, TextBoxTypeDouble, 0)
        Set This.StickerSpace = _
            TextBoxHandler.New_(StickerSpace, TextBoxTypeDouble, 0)
        Set This.QuantityOnPressSheet = _
            TextBoxHandler.New_(QuantityOnPressSheet, TextBoxTypeLong, 1)
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
    If Not CheckSize Then Exit Sub
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
    If Not CheckSize Then Exit Sub
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

Private Sub StartSettings_Click()
    This.Main.Settings
End Sub

Private Sub StartNumberSetter_Click()
    This.Main.NumberSetter
End Sub

Private Sub ExportPdf_Click()
    This.Main.ExportPdf
End Sub

Private Sub StickerRun_Click()
    This.Main.AddMarksAndSeparate Me
End Sub

Private Sub Host_SelectionChange()
    CheckControls
End Sub

'===============================================================================
' # Logic

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
        ResetAssignments.Enabled = False
        CreateImpositions.Enabled = False
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

Private Function CheckSize() As Boolean
    If This.Motifs.Count = 0 Then
        CheckSize = True
        Exit Function
    End If
    Dim Content As Shape
    Set Content = This.Motifs.Item(1).SurfaceA.Content
    Dim Answer As VbMsgBoxResult
    With ActiveSelectionRange.FirstShape
        If .SizeWidth = Content.SizeWidth _
       And .SizeHeight = Content.SizeHeight Then
            CheckSize = True
        Else
            Answer = _
                VBA.MsgBox( _
                    "Размер этого макета отличается от размера первого добавленного макета. Продолжить добавление?", _
                    vbYesNo + vbQuestion, "Несовпадение размера макета" _
                )
            If Answer = vbYes Then CheckSize = True
        End If
    End With
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
