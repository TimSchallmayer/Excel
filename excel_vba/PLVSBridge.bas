Attribute VB_Name = "PLVSBridge"
Option Explicit

Private PLVS_Ribbon As IRibbonUI
Private PLVS_ApplicationEvents As PLVSAppEvents
Private PLVS_ChartEventSinks As Collection
Private PLVS_TextTarget As Long
Private PLVS_TargetWorkbook As String
Private PLVS_TargetSheet As String
Private PLVS_TargetChart As String
Private Const PLVS_PENDING_COUNT As String = "_PLVS_ULTRA_PendingChoice_Count"
Private Const PLVS_PENDING_LABEL As String = "_PLVS_ULTRA_PendingChoice_Label"
Private Const PLVS_PENDING_ITEM_PREFIX As String = "_PLVS_ULTRA_PendingChoice_Item_"

Public Sub PLVS_RibbonOnLoad(ribbon As IRibbonUI)
    Set PLVS_Ribbon = ribbon
    Set PLVS_ApplicationEvents = New PLVSAppEvents
    Set PLVS_ApplicationEvents.ExcelApp = Application
    Set PLVS_ChartEventSinks = New Collection
    PLVS_TextTarget = 0
    PLVS_WatchCharts
End Sub

Public Sub PLVS_RecordActiveChart(ByVal activeChart As Chart)
    Dim chartObject As ChartObject
    Dim activeSheet As Worksheet
    Dim activeBook As Workbook

    Set chartObject = activeChart.Parent
    Set activeSheet = chartObject.Parent
    Set activeBook = activeSheet.Parent
    PLVS_TargetWorkbook = activeBook.FullName
    PLVS_TargetSheet = activeSheet.Name
    PLVS_TargetChart = chartObject.Name
End Sub

Private Function PLVS_PythonString(ByVal value As String) As String
    Dim escaped As String
    Dim index As Long
    Dim codePoint As Long
    Dim nextCodePoint As Long

    escaped = Chr$(39)
    index = 1
    Do While index <= Len(value)
        codePoint = AscW(Mid$(value, index, 1)) And &HFFFF&
        If codePoint >= &HD800& And codePoint <= &HDBFF& And index < Len(value) Then
            nextCodePoint = AscW(Mid$(value, index + 1, 1)) And &HFFFF&
            If nextCodePoint >= &HDC00& And nextCodePoint <= &HDFFF& Then
                codePoint = &H10000& + ((codePoint - &HD800&) * &H400&) + (nextCodePoint - &HDC00&)
                index = index + 1
            End If
        End If
        Select Case codePoint
            Case 34: escaped = escaped & Chr$(39) & " + chr(34) + " & Chr$(39)
            Case 39: escaped = escaped & "\'"
            Case 92: escaped = escaped & "\\"
            Case 9: escaped = escaped & "\t"
            Case 10: escaped = escaped & "\n"
            Case 13: escaped = escaped & "\r"
            Case 8: escaped = escaped & "\b"
            Case 12: escaped = escaped & "\f"
            Case 32 To 126
                escaped = escaped & ChrW$(codePoint)
            Case 0 To &HFFFF&
                escaped = escaped & "\u" & Right$("0000" & Hex$(codePoint), 4)
            Case Else
                escaped = escaped & Chr$(39) & " + chr(" & Trim$(Str$(codePoint)) & ") + " & Chr$(39)
        End Select
        index = index + 1
    Loop
    PLVS_PythonString = escaped & Chr$(39)
End Function

Private Function PLVS_PythonArgument(ByVal value As Variant) As String
    Select Case VarType(value)
        Case vbString
            PLVS_PythonArgument = PLVS_PythonString(CStr(value))
        Case vbBoolean
            If CBool(value) Then
                PLVS_PythonArgument = "True"
            Else
                PLVS_PythonArgument = "False"
            End If
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            PLVS_PythonArgument = Trim$(Str$(value))
        Case Else
            Err.Raise 5, "PLVS_PythonArgument", "Nicht unterstützter Python-Argumenttyp."
    End Select
End Function

Private Function PLVS_ValidPythonFunctionName(ByVal functionName As String) As Boolean
    Dim name As String
    Dim index As Long
    Dim character As String

    If Left$(functionName, 12) <> "excel_addin." Then Exit Function
    name = Mid$(functionName, 13)
    If Len(name) = 0 Then Exit Function
    For index = 1 To Len(name)
        character = Mid$(name, index, 1)
        If index = 1 Then
            If Not (character Like "[A-Za-z_]") Then Exit Function
        ElseIf Not (character Like "[A-Za-z0-9_]") Then
            Exit Function
        End If
    Next index
    PLVS_ValidPythonFunctionName = True
End Function

Public Function PLVS_BuildPythonCall(ByVal functionName As String, Optional ByVal arguments As Variant) As String
    Dim code As String
    Dim index As Long
    Dim lowerBound As Long
    Dim upperBound As Long

    If Not PLVS_ValidPythonFunctionName(functionName) Then
        Err.Raise 5, "PLVS_BuildPythonCall", "Ungültiger Python-Funktionsname."
    End If
    code = "import excel_addin; " & functionName & "("
    If Not IsMissing(arguments) Then
        If Not IsArray(arguments) Then Err.Raise 5, "PLVS_BuildPythonCall", "Python-Argumente müssen als Array übergeben werden."
        lowerBound = LBound(arguments)
        upperBound = UBound(arguments)
        For index = lowerBound To upperBound
            If index > lowerBound Then code = code & ", "
            code = code & PLVS_PythonArgument(arguments(index))
        Next index
    End If
    PLVS_BuildPythonCall = code & ")"
End Function

Private Function PLVS_ChartPythonArguments(Optional ByVal firstArgument As Variant) As Variant
    Dim activeChart As Chart
    On Error Resume Next
    Set activeChart = Application.ActiveChart
    On Error GoTo 0
    If Not activeChart Is Nothing Then PLVS_RecordActiveChart activeChart
    If Len(PLVS_TargetWorkbook) = 0 Or Len(PLVS_TargetSheet) = 0 Or Len(PLVS_TargetChart) = 0 Then
        Err.Raise 5, "PLVS_RunChartPythonCall", "Wähle zuerst ein PLVS-Diagramm aus."
    End If
    If IsMissing(firstArgument) Then
        PLVS_ChartPythonArguments = Array(PLVS_TargetWorkbook, PLVS_TargetSheet, PLVS_TargetChart)
    Else
        PLVS_ChartPythonArguments = Array(firstArgument, PLVS_TargetWorkbook, PLVS_TargetSheet, PLVS_TargetChart)
    End If
End Function

Public Function PLVS_BuildChartPythonCall(ByVal functionName As String, Optional ByVal firstArgument As Variant) As String
    PLVS_BuildChartPythonCall = PLVS_BuildPythonCall( _
        functionName, _
        PLVS_ChartPythonArguments(firstArgument))
End Function

Private Sub PLVS_ExecutePythonCode(ByVal code As String)
    RunPython code
End Sub

Private Sub PLVS_RunPythonCall(ByVal functionName As String, Optional ByVal arguments As Variant)
    If IsMissing(arguments) Then
        PLVS_ExecutePythonCode PLVS_BuildPythonCall(functionName)
    Else
        PLVS_ExecutePythonCode PLVS_BuildPythonCall(functionName, arguments)
    End If
End Sub

Public Sub PLVS_RunChartPythonCall(ByVal functionName As String, Optional ByVal firstArgument As Variant)
    Dim pythonCode As String
    If IsMissing(firstArgument) Then
        pythonCode = PLVS_BuildChartPythonCall(functionName)
    Else
        pythonCode = PLVS_BuildChartPythonCall(functionName, firstArgument)
    End If
    PLVS_ExecutePythonCode pythonCode
End Sub

Private Function PLVS_GetTargetChart() As Chart
    Dim book As Workbook
    Dim sheet As Worksheet
    If Len(PLVS_TargetWorkbook) = 0 Or Len(PLVS_TargetSheet) = 0 Or Len(PLVS_TargetChart) = 0 Then Exit Function
    For Each book In Application.Workbooks
        If StrComp(book.FullName, PLVS_TargetWorkbook, vbTextCompare) = 0 Then
            Set sheet = book.Worksheets(PLVS_TargetSheet)
            Set PLVS_GetTargetChart = sheet.ChartObjects(PLVS_TargetChart).Chart
            Exit Function
        End If
    Next book
End Function

Public Sub PLVS_RibbonGetAnalysis(control As IRibbonControl, ByRef returnedVal)
    Dim analysis As String
    Dim part As Variant
    Dim prefix As String
    Dim analysisName As String
    Dim activeChart As Chart

    On Error GoTo NoAnalysis
    analysisName = "_PLVS_ULTRA_Analysis"
    On Error Resume Next
    Set activeChart = Application.ActiveChart
    On Error GoTo NoAnalysis
    If Not activeChart Is Nothing Then
        analysisName = analysisName & "_" & CStr(activeChart.Parent.Name)
        On Error Resume Next
        analysis = CStr(Application.Evaluate(ActiveWorkbook.Names(analysisName).Name))
        If Err.Number <> 0 Then
            Err.Clear
            analysisName = "_PLVS_ULTRA_Analysis"
        End If
        On Error GoTo NoAnalysis
    End If
    If Len(analysis) = 0 Then
        analysis = CStr(Application.Evaluate(ActiveWorkbook.Names(analysisName).Name))
    End If
    Select Case control.Id
        Case "PLVSAnalysisAxes"
            prefix = "Unabhängig:"
        Case "PLVSAnalysisRelationship"
            prefix = "Zusammenhang:"
        Case "PLVSAnalysisChartFit"
            prefix = "Diagramm:"
        Case "PLVSAnalysisMetrics"
            prefix = "Messpunkte:"
        Case "PLVSAnalysisStatistics"
            prefix = "Statistik:"
        Case "PLVSAnalysisChange"
            prefix = "Veränderung:"
        Case "PLVSAnalysisAI"
            prefix = "KI-Einschätzung:"
    End Select
    For Each part In Split(analysis, " | ")
        If Left$(CStr(part), Len(prefix)) = prefix Then
            returnedVal = CStr(part)
            Exit Sub
        End If
    Next part
    If control.Id = "PLVSAnalysisAxes" Then
        returnedVal = "Diagramm erstellen, um eine Analyse zu starten."
    Else
        returnedVal = vbNullString
    End If
    Exit Sub
NoAnalysis:
    If control.Id = "PLVSAnalysisAxes" Then
        returnedVal = "Diagramm erstellen, um eine Analyse zu starten."
    Else
        returnedVal = vbNullString
    End If
    Err.Clear
End Sub

Public Sub PLVS_RibbonGetEditEnabled(control As IRibbonControl, ByRef returnedVal)
    Dim activeChart As Chart
    Dim chartType As Long

    returnedVal = False
    On Error Resume Next
    Set activeChart = PLVS_GetTargetChart()
    If activeChart Is Nothing Then Set activeChart = Application.ActiveChart
    On Error GoTo 0
    If activeChart Is Nothing Then Exit Sub
    If Left$(CStr(activeChart.Parent.Name), Len("PLVS_ULTRA_Chart_")) <> "PLVS_ULTRA_Chart_" Then Exit Sub

    chartType = activeChart.ChartType
    Select Case control.Id
        Case "PLVSLineWidth"
            returnedVal = True
        Case "PLVSSeriesColor"
            returnedVal = (chartType = -4169 Or chartType = 4 Or chartType = 65 Or chartType = 74 Or chartType = 75 Or chartType = 51 Or chartType = 57)
        Case "PLVSLineColor"
            returnedVal = (chartType = -4169 Or chartType = 4 Or chartType = 65 Or chartType = 74 Or chartType = 75)
        Case "PLVSPointSize", "PLVSTogglePoints"
            returnedVal = (chartType = -4169 Or chartType = 74 Or chartType = 75 Or chartType = 4 Or chartType = 65)
        Case "PLVSToggleConnections"
            returnedVal = (chartType = -4169 Or chartType = 74 Or chartType = 75 Or chartType = 4 Or chartType = 65)
        Case "PLVSChangeTrendline", "PLVSPolynomialDegree", "PLVSMovingAveragePeriod"
            returnedVal = (chartType = -4169 Or chartType = 74 Or chartType = 75 Or chartType = 4 Or chartType = 65 Or chartType = 51 Or chartType = 57)
        Case "PLVSFontName", "PLVSFontSize", "PLVSFontBold", "PLVSFontItalic", "PLVSFontColor"
            returnedVal = True
        Case Else
            returnedVal = True
    End Select
End Sub

Public Sub PLVS_RibbonGetEditPressed(control As IRibbonControl, ByRef returnedVal)
    Dim activeChart As Chart
    Dim chartType As Long

    returnedVal = False
    On Error Resume Next
    Set activeChart = PLVS_GetTargetChart()
    If activeChart Is Nothing Then Set activeChart = Application.ActiveChart
    If activeChart Is Nothing Then Exit Sub
    Select Case control.Id
        Case "PLVSToggleTitle"
            returnedVal = activeChart.HasTitle
        Case "PLVSToggleLegend"
            returnedVal = activeChart.HasLegend
        Case "PLVSToggleGridlines"
            returnedVal = activeChart.Axes(2).HasMajorGridlines
        Case "PLVSTogglePoints"
            returnedVal = (activeChart.SeriesCollection(1).MarkerStyle <> -4142)
        Case "PLVSToggleConnections"
            returnedVal = CBool(activeChart.SeriesCollection(1).Format.Line.Visible)
    End Select
End Sub

Public Sub PLVS_RibbonGetEditColor(control As IRibbonControl, ByRef returnedVal)
    Dim activeChart As Chart
    Dim selectedFont As Object

    returnedVal = RGB(68, 114, 196)
    On Error Resume Next
    Set activeChart = PLVS_GetTargetChart()
    If activeChart Is Nothing Then Set activeChart = Application.ActiveChart
    If activeChart Is Nothing Then Exit Sub
    Select Case control.Id
        Case "PLVSSeriesColor"
            If activeChart.ChartType = 51 Or activeChart.ChartType = 57 Then
                returnedVal = activeChart.SeriesCollection(1).Format.Fill.ForeColor.RGB
            ElseIf Not activeChart.SeriesCollection(1).Format.Line.Visible Then
                returnedVal = activeChart.SeriesCollection(1).MarkerForegroundColor
                If returnedVal < 0 Then returnedVal = activeChart.SeriesCollection(1).MarkerBackgroundColor
            Else
                returnedVal = activeChart.SeriesCollection(1).Format.Line.ForeColor.RGB
            End If
        Case "PLVSLineColor"
            returnedVal = activeChart.SeriesCollection(1).Format.Line.ForeColor.RGB
        Case "PLVSFontColor"
            Set selectedFont = PLVS_SelectedTextFont(activeChart)
            If Not selectedFont Is Nothing Then returnedVal = selectedFont.Color
    End Select
End Sub

Public Sub PLVS_RibbonGetTextStyle(control As IRibbonControl, ByRef returnedVal)
    Dim activeChart As Chart
    Dim selectedFont As Object

    returnedVal = False
    On Error Resume Next
    Set activeChart = PLVS_GetTargetChart()
    If activeChart Is Nothing Then Set activeChart = Application.ActiveChart
    If activeChart Is Nothing Then Exit Sub
    Set selectedFont = PLVS_SelectedTextFont(activeChart)
    If selectedFont Is Nothing Then Exit Sub
    If control.Id = "PLVSFontBold" Then returnedVal = selectedFont.Bold
    If control.Id = "PLVSFontItalic" Then returnedVal = selectedFont.Italic
End Sub

Private Function PLVS_SelectedTextFont(ByVal activeChart As Chart) As Object
    On Error Resume Next
    Select Case PLVS_TextTarget
        Case 0
            If activeChart.HasTitle Then Set PLVS_SelectedTextFont = activeChart.ChartTitle.Font
        Case 1
            If activeChart.Axes(1).HasTitle Then Set PLVS_SelectedTextFont = activeChart.Axes(1).AxisTitle.Font
        Case 2
            If activeChart.Axes(2).HasTitle Then Set PLVS_SelectedTextFont = activeChart.Axes(2).AxisTitle.Font
        Case 3
            If activeChart.HasLegend Then Set PLVS_SelectedTextFont = activeChart.Legend.Font
    End Select
End Function

Public Sub PLVS_RibbonGetTextTarget(control As IRibbonControl, ByRef returnedVal)
    returnedVal = PLVS_TextTarget
End Sub

Public Function PLVS_RibbonChartTypeIndex(ByVal chartType As Long) As Long
    PLVS_RibbonChartTypeIndex = 0
    Select Case chartType
        Case -4169
            PLVS_RibbonChartTypeIndex = 0
        Case 74, 75
            PLVS_RibbonChartTypeIndex = 1
        Case 4
            PLVS_RibbonChartTypeIndex = 2
        Case 65
            PLVS_RibbonChartTypeIndex = 3
        Case 57
            PLVS_RibbonChartTypeIndex = 4
        Case 51
            PLVS_RibbonChartTypeIndex = 5
    End Select
End Function

Public Sub PLVS_RibbonGetSelectedItemIndex(control As IRibbonControl, ByRef returnedVal)
    Dim activeChart As Chart
    Dim font As Object
    Dim trendline As Trendline
    Dim series As Series
    Dim widths As Variant
    Dim sizes As Variant
    Dim paletteColors As Variant
    Dim fontName As String
    Dim chartType As Long
    Dim currentValue As Double
    Dim index As Long
    Dim fontNames As Variant

    returnedVal = 0
    On Error Resume Next
    Set activeChart = PLVS_GetTargetChart()
    If activeChart Is Nothing Then Set activeChart = Application.ActiveChart
    If activeChart Is Nothing Then Exit Sub
    Select Case control.Id
        Case "PLVSTextTarget"
            returnedVal = PLVS_TextTarget
        Case "PLVSChangeChartType"
            chartType = activeChart.ChartType
            returnedVal = PLVS_RibbonChartTypeIndex(chartType)
        Case "PLVSChangeTrendline"
            returnedVal = PLVS_RibbonTrendlineSelectedIndex(activeChart)
        Case "PLVSLineWidth"
            currentValue = activeChart.SeriesCollection(1).Format.Line.Weight
            widths = Array(0.75, 1, 1.5, 2, 3)
            For index = LBound(widths) To UBound(widths)
                If Abs(currentValue - CDbl(widths(index))) < 0.01 Then returnedVal = index
            Next index
        Case "PLVSPointSize"
            currentValue = activeChart.SeriesCollection(1).MarkerSize
            sizes = Array(3, 5, 7, 9, 11)
            For index = LBound(sizes) To UBound(sizes)
                If Abs(currentValue - CDbl(sizes(index))) < 0.01 Then returnedVal = index
            Next index
        Case "PLVSSeriesColor", "PLVSLineColor", "PLVSFontColor"
            Select Case control.Id
                Case "PLVSSeriesColor"
                    If activeChart.ChartType = 51 Or activeChart.ChartType = 57 Then
                        currentValue = activeChart.SeriesCollection(1).Format.Fill.ForeColor.RGB
                    ElseIf Not activeChart.SeriesCollection(1).Format.Line.Visible Then
                        currentValue = activeChart.SeriesCollection(1).MarkerForegroundColor
                        If currentValue < 0 Then currentValue = activeChart.SeriesCollection(1).MarkerBackgroundColor
                    Else
                        currentValue = activeChart.SeriesCollection(1).Format.Line.ForeColor.RGB
                    End If
                Case "PLVSLineColor"
                    currentValue = activeChart.SeriesCollection(1).Format.Line.ForeColor.RGB
                Case "PLVSFontColor"
                    Set font = PLVS_SelectedTextFont(activeChart)
                    If font Is Nothing Then Exit Sub
                    currentValue = font.Color
            End Select
            paletteColors = Array( _
                RGB(68, 114, 196), _
                RGB(237, 125, 49), _
                RGB(165, 165, 165), _
                RGB(255, 192, 0), _
                RGB(91, 155, 213), _
                RGB(112, 173, 71), _
                RGB(38, 68, 120), _
                RGB(192, 0, 0), _
                RGB(112, 48, 160), _
                RGB(0, 0, 0))
            For index = LBound(paletteColors) To UBound(paletteColors)
                If Abs(currentValue - CDbl(paletteColors(index))) < 0.01 Then returnedVal = index
            Next index
        Case "PLVSFontName", "PLVSFontSize"
            Set font = PLVS_SelectedTextFont(activeChart)
            If font Is Nothing Then Exit Sub
            If control.Id = "PLVSFontName" Then
                fontNames = Array("Aptos", "Arial", "Calibri", "Times New Roman", "Cambria", "Verdana", "Tahoma", "Georgia", "Trebuchet MS", "Courier New", "Consolas", "Segoe UI")
                For index = LBound(fontNames) To UBound(fontNames)
                    If StrComp(CStr(font.Name), CStr(fontNames(index)), vbTextCompare) = 0 Then returnedVal = index
                Next index
            Else
                currentValue = font.Size
                sizes = Array(8, 9, 10, 11, 12, 14, 16, 18, 20, 24)
                For index = LBound(sizes) To UBound(sizes)
                    If Abs(currentValue - CDbl(sizes(index))) < 0.01 Then returnedVal = index
                Next index
            End If
    End Select
End Sub

Private Function PLVS_TrendlineDataSummary( _
    ByVal activeChart As Chart, _
    ByRef pointCount As Long, _
    ByRef distinctXCount As Long, _
    ByRef allPositiveX As Boolean, _
    ByRef allPositiveY As Boolean) As Boolean
    Dim chartType As Long
    Dim series As Series
    Dim xItems As Variant
    Dim yItems As Variant
    Dim xItem As Variant
    Dim yItem As Variant
    Dim index As Long
    Dim xValue As Double
    Dim yValue As Double
    Dim seenX As Object
    Dim key As String

    On Error GoTo InvalidData
    chartType = activeChart.ChartType
    If Not (chartType = -4169 Or chartType = 74 Or chartType = 75 Or chartType = 4 Or chartType = 65 Or chartType = 51 Or chartType = 57) Then Exit Function
    Set series = activeChart.SeriesCollection(1)
    xItems = series.XValues
    yItems = series.Values
    If Not IsArray(xItems) Or Not IsArray(yItems) Then Exit Function
    Set seenX = CreateObject("Scripting.Dictionary")
    allPositiveX = True
    allPositiveY = True
    For index = LBound(yItems) To UBound(yItems)
        yItem = yItems(index)
        If Not IsError(yItem) And IsNumeric(yItem) Then
            yValue = CDbl(yItem)
            If chartType = -4169 Or chartType = 74 Or chartType = 75 Then
                xItem = xItems(LBound(xItems) + index - LBound(yItems))
                If IsError(xItem) Or Not IsNumeric(xItem) Then GoTo InvalidData
                xValue = CDbl(xItem)
            Else
                xValue = index - LBound(yItems) + 1
            End If
            pointCount = pointCount + 1
            If yValue <= 0 Then allPositiveY = False
            If xValue <= 0 Then allPositiveX = False
            key = Format$(xValue, "0.################")
            If Not seenX.Exists(key) Then seenX.Add key, True
        End If
    Next index
    distinctXCount = seenX.Count
    PLVS_TrendlineDataSummary = (pointCount > 0)
    Exit Function
InvalidData:
    Err.Clear
    PLVS_TrendlineDataSummary = False
End Function

Private Function PLVS_TrendlineOptionCount(ByVal activeChart As Chart, ByRef optionIds As Variant, ByRef optionLabels As Variant) As Long
    Dim pointCount As Long
    Dim distinctXCount As Long
    Dim allPositiveX As Boolean
    Dim allPositiveY As Boolean
    Dim maximumDegree As Long
    Dim movingAveragePeriodCount As Long

    If Not PLVS_TrendlineDataSummary(activeChart, pointCount, distinctXCount, allPositiveX, allPositiveY) Then Exit Function
    ReDim optionIds(0 To 6)
    ReDim optionLabels(0 To 6)
    PLVS_AddTrendlineOption optionIds, optionLabels, PLVS_TrendlineOptionCount, "none", "Keine"
    If distinctXCount >= 2 Then PLVS_AddTrendlineOption optionIds, optionLabels, PLVS_TrendlineOptionCount, "linear", "Linear"
    If pointCount >= 3 And distinctXCount >= 2 And allPositiveY Then PLVS_AddTrendlineOption optionIds, optionLabels, PLVS_TrendlineOptionCount, "exponential", "Exponentiell"
    If distinctXCount >= 2 And allPositiveX Then PLVS_AddTrendlineOption optionIds, optionLabels, PLVS_TrendlineOptionCount, "logarithmic", "Logarithmisch"
    If distinctXCount >= 2 And allPositiveX And allPositiveY Then PLVS_AddTrendlineOption optionIds, optionLabels, PLVS_TrendlineOptionCount, "power", "Potenz"
    maximumDegree = pointCount - 1
    If distinctXCount - 1 < maximumDegree Then maximumDegree = distinctXCount - 1
    If maximumDegree > 6 Then maximumDegree = 6
    If maximumDegree >= 2 Then PLVS_AddTrendlineOption optionIds, optionLabels, PLVS_TrendlineOptionCount, "polynomial", "Polynomisch"
    movingAveragePeriodCount = PLVS_RibbonMovingAverageOptionCount(activeChart)
    If movingAveragePeriodCount > 0 Then PLVS_AddTrendlineOption optionIds, optionLabels, PLVS_TrendlineOptionCount, "moving_average", "Gleitender Durchschnitt"
End Function

Private Sub PLVS_AddTrendlineOption(ByRef optionIds As Variant, ByRef optionLabels As Variant, ByRef optionCount As Long, ByVal optionId As String, ByVal optionLabel As String)
    optionIds(optionCount) = optionId
    optionLabels(optionCount) = optionLabel
    optionCount = optionCount + 1
End Sub

Private Function PLVS_RibbonMovingAverageOptionCount(ByVal activeChart As Chart) As Long
    Dim pointCount As Long
    Dim distinctXCount As Long
    Dim allPositiveX As Boolean
    Dim allPositiveY As Boolean
    Dim periods As Variant
    Dim period As Variant
    Dim count As Long
    If Not PLVS_TrendlineDataSummary(activeChart, pointCount, distinctXCount, allPositiveX, allPositiveY) Then Exit Function
    periods = Array(2, 3, 4, 5, 6, 7, 10, 12, 20)
    For Each period In periods
        If CLng(period) < pointCount Then count = count + 1
    Next period
    PLVS_RibbonMovingAverageOptionCount = count
End Function

Private Function PLVS_RibbonTrendlineSelectedIndex(ByVal activeChart As Chart) As Long
    Dim series As Series
    Dim trendline As Trendline
    Dim optionIds As Variant
    Dim optionLabels As Variant
    Dim optionCount As Long
    Dim selectedId As String
    Dim index As Long
    PLVS_RibbonTrendlineSelectedIndex = 0
    optionCount = PLVS_TrendlineOptionCount(activeChart, optionIds, optionLabels)
    If optionCount = 0 Then Exit Function
    On Error GoTo NoSelection
    Set series = activeChart.SeriesCollection(1)
    If series.Trendlines.Count = 0 Then Exit Function
    Set trendline = series.Trendlines(series.Trendlines.Count)
    Select Case CLng(trendline.Type)
        Case 3: selectedId = "polynomial"
        Case 5: selectedId = "exponential"
        Case -4132: selectedId = "linear"
        Case -4133: selectedId = "logarithmic"
        Case 4: selectedId = "power"
        Case 6: selectedId = "moving_average"
    End Select
    For index = 0 To optionCount - 1
        If optionIds(index) = selectedId Then
            PLVS_RibbonTrendlineSelectedIndex = index
            Exit Function
        End If
    Next index
NoSelection:
    Err.Clear
End Function

Public Sub PLVS_RibbonGetTrendlineCount(control As IRibbonControl, ByRef returnedVal)
    Dim activeChart As Chart
    Dim optionIds As Variant
    Dim optionLabels As Variant
    On Error Resume Next
    Set activeChart = PLVS_GetTargetChart()
    If activeChart Is Nothing Then Set activeChart = Application.ActiveChart
    On Error GoTo 0
    If activeChart Is Nothing Then
        returnedVal = 0
    Else
        returnedVal = PLVS_TrendlineOptionCount(activeChart, optionIds, optionLabels)
    End If
End Sub

Public Sub PLVS_RibbonGetTrendlineLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
    Dim activeChart As Chart
    Dim optionIds As Variant
    Dim optionLabels As Variant
    Dim optionCount As Long
    On Error Resume Next
    Set activeChart = PLVS_GetTargetChart()
    If activeChart Is Nothing Then Set activeChart = Application.ActiveChart
    On Error GoTo 0
    If activeChart Is Nothing Then Exit Sub
    optionCount = PLVS_TrendlineOptionCount(activeChart, optionIds, optionLabels)
    If index >= 0 And index < optionCount Then returnedVal = optionLabels(index)
End Sub

Public Sub PLVS_RibbonGetTrendlineId(control As IRibbonControl, index As Integer, ByRef returnedVal)
    Dim activeChart As Chart
    Dim optionIds As Variant
    Dim optionLabels As Variant
    Dim optionCount As Long
    On Error Resume Next
    Set activeChart = PLVS_GetTargetChart()
    If activeChart Is Nothing Then Set activeChart = Application.ActiveChart
    On Error GoTo 0
    If activeChart Is Nothing Then Exit Sub
    optionCount = PLVS_TrendlineOptionCount(activeChart, optionIds, optionLabels)
    If index >= 0 And index < optionCount Then returnedVal = optionIds(index)
End Sub

Public Sub PLVS_RibbonGetTrendlineSelectedIndex(control As IRibbonControl, ByRef returnedVal)
    Dim activeChart As Chart
    returnedVal = 0
    On Error Resume Next
    Set activeChart = PLVS_GetTargetChart()
    If activeChart Is Nothing Then Set activeChart = Application.ActiveChart
    On Error GoTo 0
    If Not activeChart Is Nothing Then returnedVal = PLVS_RibbonTrendlineSelectedIndex(activeChart)
End Sub

Public Sub PLVS_RibbonGetTrendlineOptionVisible(control As IRibbonControl, ByRef returnedVal)
    Dim activeChart As Chart
    Dim series As Series
    Dim trendline As Trendline
    returnedVal = False
    On Error GoTo NoTrendline
    Set activeChart = PLVS_GetTargetChart()
    If activeChart Is Nothing Then Set activeChart = Application.ActiveChart
    If activeChart Is Nothing Then Exit Sub
    Set series = activeChart.SeriesCollection(1)
    If series.Trendlines.Count = 0 Then Exit Sub
    Set trendline = series.Trendlines(series.Trendlines.Count)
    If control.Id = "PLVSPolynomialDegree" Then returnedVal = (CLng(trendline.Type) = 3)
    If control.Id = "PLVSMovingAveragePeriod" Then returnedVal = (CLng(trendline.Type) = 6)
    Exit Sub
NoTrendline:
    Err.Clear
End Sub

Private Function PLVS_RibbonPolynomialDegreeMaximum(ByVal activeChart As Chart) As Long
    Dim pointCount As Long
    Dim distinctXCount As Long
    Dim allPositiveX As Boolean
    Dim allPositiveY As Boolean
    PLVS_RibbonPolynomialDegreeMaximum = 0
    If Not PLVS_TrendlineDataSummary(activeChart, pointCount, distinctXCount, allPositiveX, allPositiveY) Then Exit Function
    PLVS_RibbonPolynomialDegreeMaximum = pointCount - 1
    If distinctXCount - 1 < PLVS_RibbonPolynomialDegreeMaximum Then PLVS_RibbonPolynomialDegreeMaximum = distinctXCount - 1
    If PLVS_RibbonPolynomialDegreeMaximum > 6 Then PLVS_RibbonPolynomialDegreeMaximum = 6
End Function

Public Sub PLVS_RibbonGetPolynomialDegreeCount(control As IRibbonControl, ByRef returnedVal)
    Dim activeChart As Chart
    On Error Resume Next
    Set activeChart = PLVS_GetTargetChart()
    If activeChart Is Nothing Then Set activeChart = Application.ActiveChart
    On Error GoTo 0
    returnedVal = 0
    If Not activeChart Is Nothing Then returnedVal = Application.Max(0, PLVS_RibbonPolynomialDegreeMaximum(activeChart) - 1)
End Sub

Public Sub PLVS_RibbonGetPolynomialDegreeLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
    returnedVal = "Grad " & CStr(index + 2)
End Sub

Public Sub PLVS_RibbonGetPolynomialDegreeId(control As IRibbonControl, index As Integer, ByRef returnedVal)
    returnedVal = "PLVSPolynomialDegree" & CStr(index + 2)
End Sub

Public Sub PLVS_RibbonGetPolynomialDegreeSelectedIndex(control As IRibbonControl, ByRef returnedVal)
    Dim activeChart As Chart
    Dim trendline As Trendline
    returnedVal = 0
    On Error GoTo NoTrendline
    Set activeChart = PLVS_GetTargetChart()
    If activeChart Is Nothing Then Set activeChart = Application.ActiveChart
    Set trendline = activeChart.SeriesCollection(1).Trendlines(activeChart.SeriesCollection(1).Trendlines.Count)
    If CLng(trendline.Type) = 3 Then returnedVal = CLng(trendline.Order) - 2
NoTrendline:
    Err.Clear
End Sub

Private Function PLVS_RibbonMovingAveragePeriodAt(ByVal activeChart As Chart, ByVal selectionIndex As Long) As Long
    Dim pointCount As Long
    Dim distinctXCount As Long
    Dim allPositiveX As Boolean
    Dim allPositiveY As Boolean
    Dim periods As Variant
    Dim period As Variant
    Dim index As Long
    If Not PLVS_TrendlineDataSummary(activeChart, pointCount, distinctXCount, allPositiveX, allPositiveY) Then Exit Function
    periods = Array(2, 3, 4, 5, 6, 7, 10, 12, 20)
    For Each period In periods
        If CLng(period) < pointCount Then
            If index = selectionIndex Then
                PLVS_RibbonMovingAveragePeriodAt = CLng(period)
                Exit Function
            End If
            index = index + 1
        End If
    Next period
End Function

Public Sub PLVS_RibbonGetMovingAveragePeriodCount(control As IRibbonControl, ByRef returnedVal)
    Dim activeChart As Chart
    On Error Resume Next
    Set activeChart = PLVS_GetTargetChart()
    If activeChart Is Nothing Then Set activeChart = Application.ActiveChart
    On Error GoTo 0
    returnedVal = 0
    If Not activeChart Is Nothing Then returnedVal = PLVS_RibbonMovingAverageOptionCount(activeChart)
End Sub

Public Sub PLVS_RibbonGetMovingAveragePeriodLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
    Dim activeChart As Chart
    On Error Resume Next
    Set activeChart = PLVS_GetTargetChart()
    If activeChart Is Nothing Then Set activeChart = Application.ActiveChart
    On Error GoTo 0
    If Not activeChart Is Nothing Then returnedVal = CStr(PLVS_RibbonMovingAveragePeriodAt(activeChart, index))
End Sub

Public Sub PLVS_RibbonGetMovingAveragePeriodId(control As IRibbonControl, index As Integer, ByRef returnedVal)
    Dim activeChart As Chart
    On Error Resume Next
    Set activeChart = PLVS_GetTargetChart()
    If activeChart Is Nothing Then Set activeChart = Application.ActiveChart
    On Error GoTo 0
    If Not activeChart Is Nothing Then returnedVal = "PLVSMovingAveragePeriod" & CStr(PLVS_RibbonMovingAveragePeriodAt(activeChart, index))
End Sub

Public Sub PLVS_RibbonGetMovingAveragePeriodSelectedIndex(control As IRibbonControl, ByRef returnedVal)
    Dim activeChart As Chart
    Dim series As Series
    Dim trendline As Trendline
    Dim count As Long
    Dim index As Long
    returnedVal = 0
    On Error GoTo NoTrendline
    Set activeChart = PLVS_GetTargetChart()
    If activeChart Is Nothing Then Set activeChart = Application.ActiveChart
    Set series = activeChart.SeriesCollection(1)
    count = CLng(series.Trendlines.Count)
    If count = 0 Then Exit Sub
    Set trendline = series.Trendlines(count)
    If CLng(trendline.Type) <> 6 Then Exit Sub
    Do While PLVS_RibbonMovingAveragePeriodAt(activeChart, index) > 0
        If PLVS_RibbonMovingAveragePeriodAt(activeChart, index) = CLng(trendline.Period) Then
            returnedVal = index
            Exit Do
        End If
        index = index + 1
    Loop
NoTrendline:
    Err.Clear
End Sub

Public Sub PLVS_RefreshAnalysis()
    If Not PLVS_Ribbon Is Nothing Then
        PLVS_Ribbon.InvalidateControl "PLVSPendingChoice"
        PLVS_Ribbon.InvalidateControl "PLVSAnalysisAxes"
        PLVS_Ribbon.InvalidateControl "PLVSAnalysisRelationship"
        PLVS_Ribbon.InvalidateControl "PLVSAnalysisChartFit"
        PLVS_Ribbon.InvalidateControl "PLVSAnalysisMetrics"
        PLVS_Ribbon.InvalidateControl "PLVSAnalysisStatistics"
        PLVS_Ribbon.InvalidateControl "PLVSAnalysisChange"
        PLVS_Ribbon.InvalidateControl "PLVSAnalysisAI"
        PLVS_Ribbon.InvalidateControl "PLVSSwapAxes"
        PLVS_Ribbon.InvalidateControl "PLVSChangeChartType"
        PLVS_Ribbon.InvalidateControl "PLVSSeriesColor"
        PLVS_Ribbon.InvalidateControl "PLVSLineColor"
        PLVS_Ribbon.InvalidateControl "PLVSLineWidth"
        PLVS_Ribbon.InvalidateControl "PLVSPointSize"
        PLVS_Ribbon.InvalidateControl "PLVSTogglePoints"
        PLVS_Ribbon.InvalidateControl "PLVSToggleConnections"
        PLVS_Ribbon.InvalidateControl "PLVSChangeTrendline"
        PLVS_Ribbon.InvalidateControl "PLVSPolynomialDegree"
        PLVS_Ribbon.InvalidateControl "PLVSMovingAveragePeriod"
        PLVS_Ribbon.InvalidateControl "PLVSEditTitle"
        PLVS_Ribbon.InvalidateControl "PLVSEditXAxisTitle"
        PLVS_Ribbon.InvalidateControl "PLVSEditYAxisTitle"
        PLVS_Ribbon.InvalidateControl "PLVSToggleLegend"
        PLVS_Ribbon.InvalidateControl "PLVSToggleGridlines"
        PLVS_Ribbon.InvalidateControl "PLVSToggleTitle"
        PLVS_Ribbon.InvalidateControl "PLVSFontBold"
        PLVS_Ribbon.InvalidateControl "PLVSFontItalic"
        PLVS_Ribbon.InvalidateControl "PLVSFontColor"
        PLVS_Ribbon.InvalidateControl "PLVSFontName"
        PLVS_Ribbon.InvalidateControl "PLVSFontSize"
        PLVS_Ribbon.InvalidateControl "PLVSTextTarget"
    End If
End Sub

Private Function PLVS_PendingNameValue(ByVal shortName As String) As String
    Dim workbook As Workbook
    Dim item As Name
    Dim itemName As String
    Dim separator As Long

    On Error GoTo NoValue
    Set workbook = Application.ActiveWorkbook
    If workbook Is Nothing Then Exit Function
    For Each item In workbook.Names
        itemName = item.Name
        separator = InStrRev(itemName, "!")
        If separator > 0 Then itemName = Mid$(itemName, separator + 1)
        If StrComp(itemName, shortName, vbTextCompare) = 0 Then
            PLVS_PendingNameValue = CStr(Application.Evaluate(item.RefersTo))
            Exit Function
        End If
    Next item
NoValue:
    Err.Clear
End Function

Public Sub PLVS_RibbonGetPendingVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = (Val(PLVS_PendingNameValue(PLVS_PENDING_COUNT)) > 0)
End Sub

Public Sub PLVS_RibbonGetPendingItemCount(control As IRibbonControl, ByRef returnedVal)
    returnedVal = CLng(Val(PLVS_PendingNameValue(PLVS_PENDING_COUNT)))
End Sub

Public Sub PLVS_RibbonGetPendingItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
    returnedVal = PLVS_PendingNameValue(PLVS_PENDING_ITEM_PREFIX & CStr(index))
End Sub

Public Sub PLVS_RibbonGetPendingItemId(control As IRibbonControl, index As Integer, ByRef returnedVal)
    returnedVal = "PLVSPendingChoice" & CStr(index)
End Sub

Public Sub PLVS_WatchCharts()
    Dim book As Workbook
    Dim sheet As Worksheet
    Dim chartObject As ChartObject
    Dim chartEvents As PLVSChartEvents
    Dim chartKey As String

    Set PLVS_ChartEventSinks = New Collection
    For Each book In Application.Workbooks
        For Each sheet In book.Worksheets
            For Each chartObject In sheet.ChartObjects
                chartKey = book.FullName & "|" & sheet.CodeName & "|" & chartObject.Name
                Set chartEvents = New PLVSChartEvents
                chartEvents.Connect chartObject.Chart
                PLVS_ChartEventSinks.Add chartEvents, chartKey
            Next chartObject
        Next sheet
    Next book
End Sub

Public Sub PLVS_ULTRA_CreateChart()
    PLVS_RunPythonCall "excel_addin.create_chart"
    PLVS_WatchCharts
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_ULTRA_TestAI()
    PLVS_RunPythonCall "excel_addin.test_ai"
End Sub

Public Sub PLVS_ULTRA_AISettings()
    PLVS_RunPythonCall "excel_addin.open_ai_settings"
End Sub

Public Sub PLVS_ULTRA_Help()
    PLVS_RunPythonCall "excel_addin.show_help"
End Sub

Public Sub PLVS_RibbonCreateChart(control As IRibbonControl)
    PLVS_RunPythonCall "excel_addin.create_chart"
    PLVS_WatchCharts
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonChoosePendingChoice(control As IRibbonControl, selectedId As String, index As Integer)
    PLVS_RunPythonCall "excel_addin.choose_pending_choice", Array(CLng(index))
    PLVS_WatchCharts
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonSwapAxes(control As IRibbonControl)
    PLVS_RunChartPythonCall "excel_addin.swap_axes"
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonChangeChartType(control As IRibbonControl, selectedId As String, index As Integer)
    PLVS_RunChartPythonCall "excel_addin.change_chart_type", index
    PLVS_RefreshAnalysis
End Sub

Private Function PLVS_RibbonColorFromIndex(ByVal index As Integer) As Long
    Select Case index
        Case 0: PLVS_RibbonColorFromIndex = RGB(68, 114, 196)
        Case 1: PLVS_RibbonColorFromIndex = RGB(237, 125, 49)
        Case 2: PLVS_RibbonColorFromIndex = RGB(165, 165, 165)
        Case 3: PLVS_RibbonColorFromIndex = RGB(255, 192, 0)
        Case 4: PLVS_RibbonColorFromIndex = RGB(91, 155, 213)
        Case 5: PLVS_RibbonColorFromIndex = RGB(112, 173, 71)
        Case 6: PLVS_RibbonColorFromIndex = RGB(38, 68, 120)
        Case 7: PLVS_RibbonColorFromIndex = RGB(192, 0, 0)
        Case 8: PLVS_RibbonColorFromIndex = RGB(112, 48, 160)
        Case 9: PLVS_RibbonColorFromIndex = RGB(0, 0, 0)
        Case Else: Err.Raise 5, "PLVS_RibbonColorFromIndex", "Ungültige Farbauswahl."
    End Select
End Function

Public Sub PLVS_RibbonSeriesColor(control As IRibbonControl, selectedId As String, index As Integer)
    PLVS_RunChartPythonCall "excel_addin.set_series_color", PLVS_RibbonColorFromIndex(index)
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonLineColor(control As IRibbonControl, selectedId As String, index As Integer)
    PLVS_RunChartPythonCall "excel_addin.set_line_color", PLVS_RibbonColorFromIndex(index)
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonLineWidth(control As IRibbonControl, selectedId As String, index As Integer)
    PLVS_RunChartPythonCall "excel_addin.set_line_width", index
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonPointSize(control As IRibbonControl, selectedId As String, index As Integer)
    PLVS_RunChartPythonCall "excel_addin.set_point_size", index
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonTogglePoints(control As IRibbonControl, pressed As Boolean)
    PLVS_RunChartPythonCall "excel_addin.toggle_points"
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonToggleConnections(control As IRibbonControl, pressed As Boolean)
    PLVS_RunChartPythonCall "excel_addin.set_connections", pressed
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonChangeTrendline(control As IRibbonControl, selectedId As String, index As Integer)
    PLVS_RunChartPythonCall "excel_addin.change_trendline", index
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonChangePolynomialDegree(control As IRibbonControl, selectedId As String, index As Integer)
    PLVS_RunChartPythonCall "excel_addin.change_polynomial_degree", index
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonChangeMovingAveragePeriod(control As IRibbonControl, selectedId As String, index As Integer)
    PLVS_RunChartPythonCall "excel_addin.change_moving_average_period", index
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonToggleTitle(control As IRibbonControl, pressed As Boolean)
    PLVS_RunChartPythonCall "excel_addin.toggle_title"
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonEditTitle(control As IRibbonControl)
    PLVS_RunChartPythonCall "excel_addin.edit_chart_title"
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonEditXAxisTitle(control As IRibbonControl)
    PLVS_RunChartPythonCall "excel_addin.edit_x_axis_title"
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonEditYAxisTitle(control As IRibbonControl)
    PLVS_RunChartPythonCall "excel_addin.edit_y_axis_title"
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonToggleLegend(control As IRibbonControl, pressed As Boolean)
    PLVS_RunChartPythonCall "excel_addin.set_legend", pressed
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonToggleGridlines(control As IRibbonControl, pressed As Boolean)
    PLVS_RunChartPythonCall "excel_addin.toggle_gridlines"
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonTextTarget(control As IRibbonControl, selectedId As String, index As Integer)
    PLVS_TextTarget = index
    PLVS_RunChartPythonCall "excel_addin.set_text_target", index
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonFontName(control As IRibbonControl, selectedId As String, index As Integer)
    PLVS_RunChartPythonCall "excel_addin.set_text_font_name", index
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonFontSize(control As IRibbonControl, selectedId As String, index As Integer)
    PLVS_RunChartPythonCall "excel_addin.set_text_font_size", index
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonFontColor(control As IRibbonControl, selectedId As String, index As Integer)
    PLVS_RunChartPythonCall "excel_addin.set_text_color", PLVS_RibbonColorFromIndex(index)
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonToggleBold(control As IRibbonControl, pressed As Boolean)
    PLVS_RunChartPythonCall "excel_addin.toggle_text_bold"
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonToggleItalic(control As IRibbonControl, pressed As Boolean)
    PLVS_RunChartPythonCall "excel_addin.toggle_text_italic"
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonTestAI(control As IRibbonControl)
    PLVS_RunPythonCall "excel_addin.test_ai"
End Sub

Public Sub PLVS_RibbonAISettings(control As IRibbonControl)
    PLVS_RunPythonCall "excel_addin.open_ai_settings"
End Sub

Public Sub PLVS_RibbonHelp(control As IRibbonControl)
    PLVS_RunPythonCall "excel_addin.show_help"
End Sub
