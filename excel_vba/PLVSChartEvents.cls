VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PLVSChartEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents PLVS_Chart As Chart

Public Sub Connect(ByVal chart As Chart)
    Set PLVS_Chart = chart
End Sub

Private Sub PLVS_Chart_Activate()
    PLVS_RefreshAnalysis
End Sub

Private Sub PLVS_Chart_Deactivate()
    PLVS_RefreshAnalysis
End Sub
