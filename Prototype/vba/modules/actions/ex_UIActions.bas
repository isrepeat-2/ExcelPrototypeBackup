Attribute VB_Name = "ex_UIActions"
Option Explicit

' UI entrypoints layer: keeps user-triggered callbacks in actions/*
' and delegates work to domain/config modules.

Public Sub m_DeleteResultSheets()
    ex_SheetStylesXmlProvider.m_DeleteResultSheets
End Sub

Public Sub m_SwitchMode_OnClick()
    ex_Settings.m_SwitchMode_OnClick
End Sub

Public Sub m_OnProfileChanged()
    ex_ConfigProfilesManager.m_OnProfileChanged
End Sub

Public Sub m_OnModeChanged()
    ex_ConfigProfilesManager.m_OnModeChanged
End Sub

Public Sub m_HelloWorld()
    ex_Startup.m_HelloWorld
End Sub

Public Sub m_OutputPanelStartSearch()
    Dim ws As Worksheet
    Dim searchKey As String
    Dim outputStyle As ex_SheetStylesXmlProvider.t_OutputSheetStyle
    Dim configKey As String

    Set ws = ActiveSheet
    If ws Is Nothing Then
        MsgBox "Active sheet is not available for output panel search.", vbExclamation
        Exit Sub
    End If

    searchKey = ex_OutputPanel.m_ReadSearchValue(ws)
    If Len(searchKey) = 0 Then
        MsgBox "Введите значение ключа в панели поиска.", vbExclamation
        Exit Sub
    End If

    configKey = "Context.PersonValue"
    If ex_SheetStylesXmlProvider.m_GetOutputSheetStyle(outputStyle, ThisWorkbook) Then
        If Len(Trim$(outputStyle.PanelInputConfigKey)) > 0 Then
            configKey = Trim$(outputStyle.PanelInputConfigKey)
        End If
    End If

    ex_ConfigProvider.m_SetConfigValue configKey, searchKey, True
    ex_PersonTimeline.m_ShowPersonTimeline searchKey
End Sub
