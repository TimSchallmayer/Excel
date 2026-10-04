Attribute VB_Name = "PLVSBridge"
Option Explicit

Public Sub PLVS_ULTRA_Analyze()
    RunPython "import excel_addin; excel_addin.analyze_table()"
End Sub

Public Sub PLVS_ULTRA_ShowAnalysis()
    RunPython "import excel_addin; excel_addin.show_analysis()"
End Sub

Public Sub PLVS_ULTRA_CreateChart()
    RunPython "import excel_addin; excel_addin.create_chart()"
End Sub

Public Sub PLVS_ULTRA_AISettings()
    RunPython "import excel_addin; excel_addin.open_ai_settings()"
End Sub

Public Sub PLVS_ULTRA_Help()
    RunPython "import excel_addin; excel_addin.show_help()"
End Sub

Public Sub PLVS_ULTRA_EnsureDashboard()
    RunPython "import excel_addin; excel_addin.ensure_dashboard()"
End Sub

Public Sub PLVS_RibbonAnalyze(control As IRibbonControl)
    RunPython "import excel_addin; excel_addin.analyze_table()"
End Sub

Public Sub PLVS_RibbonCreateChart(control As IRibbonControl)
    RunPython "import excel_addin; excel_addin.create_chart()"
End Sub

Public Sub PLVS_RibbonTestAI(control As IRibbonControl)
    RunPython "import excel_addin; excel_addin.test_ai()"
End Sub

Public Sub PLVS_RibbonAISettings(control As IRibbonControl)
    RunPython "import excel_addin; excel_addin.open_ai_settings()"
End Sub

Public Sub PLVS_RibbonShowAnalysis(control As IRibbonControl)
    RunPython "import excel_addin; excel_addin.show_analysis()"
End Sub

Public Sub PLVS_RibbonHelp(control As IRibbonControl)
    RunPython "import excel_addin; excel_addin.show_help()"
End Sub
