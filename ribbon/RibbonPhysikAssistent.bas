Attribute VB_Name = "RibbonPhysikAssistent"
Option Explicit

Public Sub CreateChart(control As IRibbonControl)
    RunPython "import excel_addin; excel_addin.create_chart()"
End Sub

Public Sub AnalyzeTable(control As IRibbonControl)
    RunPython "import excel_addin; excel_addin.analyze_table()"
End Sub

Public Sub TestAI(control As IRibbonControl)
    RunPython "import excel_addin; excel_addin.test_ai()"
End Sub

Public Sub OpenAISettings(control As IRibbonControl)
    RunPython "import excel_addin; excel_addin.open_ai_settings()"
End Sub

Public Sub ShowAnalysis(control As IRibbonControl)
    RunPython "import excel_addin; excel_addin.show_analysis()"
End Sub

Public Sub OpenSettings(control As IRibbonControl)
    RunPython "import excel_addin; excel_addin.open_settings()"
End Sub
