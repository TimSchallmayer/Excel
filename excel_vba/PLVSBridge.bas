Attribute VB_Name = "PLVSBridge"
Option Explicit

Private PLVS_Ribbon As IRibbonUI

Public Sub PLVS_RibbonOnLoad(ribbon As IRibbonUI)
    Set PLVS_Ribbon = ribbon
End Sub

Public Sub PLVS_RibbonGetAnalysis(control As IRibbonControl, ByRef returnedVal)
    Dim analysis As String
    Dim part As Variant
    Dim prefix As String

    On Error GoTo NoAnalysis
    analysis = CStr(Application.Evaluate(ActiveWorkbook.Names("_PLVS_ULTRA_Analysis").RefersTo))
    Select Case control.Id
        Case "PLVSAnalysisAxes"
            prefix = "Unabhängig:"
        Case "PLVSAnalysisChartFit"
            prefix = "Diagramm:"
        Case "PLVSAnalysisMetrics"
            prefix = "Messpunkte:"
    End Select
    For Each part In Split(analysis, " | ")
        If Left$(CStr(part), Len(prefix)) = prefix Then
            returnedVal = CStr(part)
            Exit Sub
        End If
    Next part
    If control.Id = "PLVSAnalysisAxes" Then
        returnedVal = "Datenquelle auswählen oder Diagramm erstellen."
    Else
        returnedVal = vbNullString
    End If
    Exit Sub
NoAnalysis:
    If control.Id = "PLVSAnalysisAxes" Then
        returnedVal = "Datenquelle auswählen oder Diagramm erstellen."
    Else
        returnedVal = vbNullString
    End If
    Err.Clear
End Sub

Private Sub PLVS_RefreshAnalysis()
    If Not PLVS_Ribbon Is Nothing Then
        PLVS_Ribbon.InvalidateControl "PLVSAnalysisAxes"
        PLVS_Ribbon.InvalidateControl "PLVSAnalysisChartFit"
        PLVS_Ribbon.InvalidateControl "PLVSAnalysisMetrics"
    End If
End Sub

Public Sub PLVS_ULTRA_CreateChart()
    RunPython "import excel_addin; excel_addin.create_chart()"
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_ULTRA_SelectDataSource()
    RunPython "import excel_addin; excel_addin.select_data_source()"
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
    PLVS_RefreshAnalysis
End Sub

Public Sub PLVS_RibbonSelectDataSource(control As IRibbonControl)
    RunPython "import excel_addin; excel_addin.select_data_source()"
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
