Attribute VB_Name = "PLVSBridge"
Option Explicit

Private PLVS_Ribbon As IRibbonUI
Private PLVS_ApplicationEvents As PLVSAppEvents
Private PLVS_ChartEventSinks As Collection

Public Sub PLVS_RibbonOnLoad(ribbon As IRibbonUI)
    Set PLVS_Ribbon = ribbon
    Set PLVS_ApplicationEvents = New PLVSAppEvents
    Set PLVS_ApplicationEvents.ExcelApp = Application
    Set PLVS_ChartEventSinks = New Collection
    PLVS_WatchCharts
End Sub

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

Public Sub PLVS_RefreshAnalysis()
    If Not PLVS_Ribbon Is Nothing Then
        PLVS_Ribbon.InvalidateControl "PLVSAnalysisAxes"
        PLVS_Ribbon.InvalidateControl "PLVSAnalysisRelationship"
        PLVS_Ribbon.InvalidateControl "PLVSAnalysisChartFit"
        PLVS_Ribbon.InvalidateControl "PLVSAnalysisMetrics"
        PLVS_Ribbon.InvalidateControl "PLVSAnalysisStatistics"
        PLVS_Ribbon.InvalidateControl "PLVSAnalysisChange"
        PLVS_Ribbon.InvalidateControl "PLVSAnalysisAI"
    End If
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
    RunPython "import excel_addin; excel_addin.create_chart()"
    PLVS_WatchCharts
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_ULTRA_TestAI()
    RunPython "import excel_addin; excel_addin.test_ai()"
End Sub

Public Sub PLVS_ULTRA_AISettings()
    RunPython "import excel_addin; excel_addin.open_ai_settings()"
End Sub

Public Sub PLVS_ULTRA_Help()
    RunPython "import excel_addin; excel_addin.show_help()"
End Sub

Public Sub PLVS_RibbonCreateChart(control As IRibbonControl)
    RunPython "import excel_addin; excel_addin.create_chart()"
    PLVS_WatchCharts
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonTestAI(control As IRibbonControl)
    RunPython "import excel_addin; excel_addin.test_ai()"
End Sub

Public Sub PLVS_RibbonAISettings(control As IRibbonControl)
    RunPython "import excel_addin; excel_addin.open_ai_settings()"
End Sub

Public Sub PLVS_RibbonHelp(control As IRibbonControl)
    RunPython "import excel_addin; excel_addin.show_help()"
End Sub
