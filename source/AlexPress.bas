Attribute VB_Name = "AlexPress"
'===============================================================================
'   ������          : AlexPress
'   ������          : 2023.01.23
'   �����           : https://vk.com/elvin_macro/Bleeds
'                     https://github.com/elvin-nsk/Bleeds
'   �����           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = True

Public Const APP_NAME As String = "AlexPress"

'===============================================================================

Sub Start()
    With New MainView
        .Show vbModeless
    End With
End Sub

'===============================================================================
