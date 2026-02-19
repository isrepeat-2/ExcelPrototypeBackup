Attribute VB_Name = "ex_OutputPanel"
Option Explicit

Private Const PANEL_INPUT_NAME As String = "outPanelInputCell"
Private Const PANEL_BUTTON_PREFIX As String = "btnOutPanelSearch_"

Public Sub m_RenderForSheet(ByVal ws As Worksheet, ByRef style As ex_SheetStylesXmlProvider.t_OutputSheetStyle)
    Dim topRow As Long
    Dim startCol As Long
    Dim rightCol As Long
    Dim bottomRow As Long
    Dim dataLastCol As Long
    Dim panelRange As Range
    Dim titleRange As Range
    Dim labelCell As Range
    Dim inputRange As Range
    Dim inputAnchor As Range
    Dim buttonRange As Range
    Dim buttonShape As Shape
    Dim buttonName As String
    Dim buttonMacro As String
    Dim inputStartCol As Long
    Dim inputEndCol As Long
    Dim buttonStartCol As Long
    Dim currentValue As String

    If ws Is Nothing Then Exit Sub
    If Not style.HasControlPanel Then Exit Sub

    topRow = style.PanelTopRow
    If topRow < 1 Then topRow = 1

    dataLastCol = mp_GetLastUsedColumn(ws)
    startCol = dataLastCol + style.PanelOffsetColumns
    If startCol < style.PanelMinStartColumn Then startCol = style.PanelMinStartColumn
    If startCol < 1 Then startCol = 1

    rightCol = startCol + style.PanelWidthColumns - 1
    bottomRow = topRow + style.PanelHeightRows - 1
    If rightCol <= startCol Then rightCol = startCol + 1
    If bottomRow <= topRow Then bottomRow = topRow + 2

    Set panelRange = ws.Range(ws.Cells(topRow, startCol), ws.Cells(bottomRow, rightCol))
    panelRange.Interior.Pattern = xlSolid
    panelRange.Interior.Color = style.PanelBackColor
    panelRange.Font.Color = style.PanelLabelColor
    panelRange.Font.Name = style.FontName
    panelRange.Font.Size = style.FontSize
    panelRange.Borders.LineStyle = xlContinuous
    panelRange.Borders.Color = style.PanelBorderColor
    panelRange.Borders.Weight = xlThin

    Set titleRange = ws.Range(ws.Cells(topRow, startCol), ws.Cells(topRow, rightCol))
    titleRange.UnMerge
    titleRange.Merge
    titleRange.Value = style.PanelTitle
    titleRange.Font.Bold = True
    titleRange.Font.Color = style.PanelTitleColor
    titleRange.HorizontalAlignment = xlLeft
    titleRange.VerticalAlignment = xlCenter

    Set labelCell = ws.Cells(topRow + 1, startCol)
    labelCell.Value = style.PanelInputLabel
    labelCell.Font.Bold = True
    labelCell.Font.Color = style.PanelLabelColor
    labelCell.HorizontalAlignment = xlLeft
    labelCell.VerticalAlignment = xlCenter

    inputStartCol = startCol + 1
    inputEndCol = rightCol - 2
    If inputEndCol < inputStartCol Then inputEndCol = inputStartCol
    buttonStartCol = inputEndCol + 1
    If buttonStartCol > rightCol Then buttonStartCol = rightCol

    Set inputRange = ws.Range(ws.Cells(topRow + 1, inputStartCol), ws.Cells(bottomRow, inputEndCol))
    inputRange.UnMerge
    inputRange.Merge
    inputRange.Interior.Pattern = xlSolid
    inputRange.Interior.Color = style.PanelInputBackColor
    inputRange.Font.Color = style.PanelInputFontColor
    inputRange.HorizontalAlignment = xlLeft
    inputRange.VerticalAlignment = xlCenter
    inputRange.NumberFormat = "@"

    Set inputAnchor = inputRange.Cells(1, 1)
    currentValue = Trim$(CStr(inputAnchor.Value))
    If Len(currentValue) = 0 Then
        inputAnchor.Value = ex_ConfigProvider.m_GetConfigValue(style.PanelInputConfigKey, vbNullString)
    End If

    mp_SetPanelInputName ws, inputAnchor

    buttonName = mp_GetButtonName(ws)
    mp_DeleteShapeIfExists ws, buttonName

    Set buttonRange = ws.Range(ws.Cells(topRow + 1, buttonStartCol), ws.Cells(bottomRow, rightCol))
    Set buttonShape = ws.Shapes.AddShape(msoShapeRoundedRectangle, buttonRange.Left + 1, buttonRange.Top + 1, buttonRange.Width - 2, buttonRange.Height - 2)
    buttonShape.Name = buttonName
    buttonShape.TextFrame.Characters.Text = style.PanelButtonCaption
    buttonShape.Fill.ForeColor.RGB = style.PanelButtonBackColor
    buttonShape.Line.ForeColor.RGB = style.PanelButtonBorderColor
    buttonShape.Line.Weight = 1
    buttonShape.TextFrame.Characters.Font.Bold = True
    buttonShape.TextFrame.Characters.Font.Color = style.PanelButtonTextColor
    buttonShape.TextFrame.Characters.Font.Name = style.FontName
    buttonShape.TextFrame.Characters.Font.Size = style.FontSize
    buttonShape.Placement = xlFreeFloating

    buttonMacro = Trim$(style.PanelButtonMacro)
    If Len(buttonMacro) = 0 Then
        buttonMacro = "ex_UIActions.m_OutputPanelStartSearch"
    End If
    buttonShape.OnAction = "'" & ThisWorkbook.Name & "'!" & buttonMacro
End Sub

Public Function m_ReadSearchValue(ByVal ws As Worksheet) As String
    Dim inputCell As Range
    Set inputCell = mp_GetPanelInputCell(ws)
    If inputCell Is Nothing Then Exit Function
    m_ReadSearchValue = Trim$(CStr(inputCell.Value))
End Function

Private Function mp_GetLastUsedColumn(ByVal ws As Worksheet) As Long
    Dim lastCell As Range

    If ws Is Nothing Then
        mp_GetLastUsedColumn = 1
        Exit Function
    End If

    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If lastCell Is Nothing Then
        mp_GetLastUsedColumn = 1
    Else
        mp_GetLastUsedColumn = lastCell.Column
    End If
End Function

Private Sub mp_SetPanelInputName(ByVal ws As Worksheet, ByVal inputCell As Range)
    If ws Is Nothing Then Exit Sub
    If inputCell Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Names(PANEL_INPUT_NAME).Delete
    On Error GoTo 0

    On Error Resume Next
    ws.Names.Add Name:=PANEL_INPUT_NAME, RefersTo:="=" & inputCell.Address(True, True, xlA1, True)
    On Error GoTo 0
End Sub

Private Function mp_GetPanelInputCell(ByVal ws As Worksheet) As Range
    Dim inputName As Name
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set inputName = ws.Names(PANEL_INPUT_NAME)
    On Error GoTo 0
    If inputName Is Nothing Then Exit Function

    On Error Resume Next
    Set mp_GetPanelInputCell = inputName.RefersToRange
    On Error GoTo 0
End Function

Private Function mp_GetButtonName(ByVal ws As Worksheet) As String
    If ws Is Nothing Then
        mp_GetButtonName = PANEL_BUTTON_PREFIX
        Exit Function
    End If
    mp_GetButtonName = PANEL_BUTTON_PREFIX & ws.CodeName
End Function

Private Sub mp_DeleteShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String)
    Dim shp As Shape
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If shp Is Nothing Then Exit Sub
    On Error Resume Next
    shp.Delete
    On Error GoTo 0
End Sub
