VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PLVSAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ExcelApp As Application

Private Sub ExcelApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    PLVS_RefreshAnalysis
End Sub

Private Sub ExcelApp_SheetActivate(ByVal Sh As Object)
    PLVS_WatchCharts
    PLVS_RefreshAnalysis
End Sub

Private Sub ExcelApp_WorkbookOpen(ByVal Wb As Workbook)
    PLVS_WatchCharts
    PLVS_RefreshAnalysis
End Sub

Private Sub ExcelApp_WorkbookActivate(ByVal Wb As Workbook)
    PLVS_WatchCharts
    PLVS_RefreshAnalysis
End Sub
