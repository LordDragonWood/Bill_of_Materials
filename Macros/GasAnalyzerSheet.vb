Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Option Explicit

    Dim GAAssemblyPN As String
    Dim GATrapPN As String

Private Sub Worksheet_Activate()
'Created for Pason by Dragon Wood (August 2015)
'Automatically sets the Required equipment to "Yes" and sets the default value for each. These can of course be changed.
  
  Application.EnableEvents = False
  
  Call ShowLegend
  
    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt1 Then
        GAAssemblyPN = "GSA001"
        GATrapPN = "GSA002"
    ElseIf ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt2 Then
        GAAssemblyPN = "GSA006"
        GATrapPN = "GSA010"
    End If

  
    'Only choose the 110V items if 110V is selected.
    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt1 Then
        With ActiveWorkbook.Worksheets("Gas Analyzer")
            .Range("C2").Value = GAAssemblyPN
            .Range("C3").Value = GATrapPN
            .Range("C26").Value = ADRPowerBarPN1
        End With
    End If

    'Only choose the 220V items if 220V is selected.
    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt2 Then
        With ActiveWorkbook.Worksheets("Gas Analyzer")
            .Range("C2").Value = GAAssemblyPN
            .Range("C3").Value = GATrapPN
            .Range("C26").Value = ADRPowerBarPN2
        End With
    End If

    'Lock all UJB related cells & Unlock all EDR related cells.
    If ActiveWorkbook.Worksheets("System Selection").UJBBox.Value = strNo Then
        With ActiveWorkbook.Worksheets("Gas Analyzer")
            .Range("A16:A23").Locked = True
            .Range("A16:A23").Value = strNo
            .Range("D16:D23").Locked = True
            .Range("D16:D23").ClearContents
            .Range("A27:A28").Locked = True
            .Range("A27:A28").Value = strNo
            .Range("D27:D28").Locked = True
            .Range("D27:D28").ClearContents
        End With
    End If

    'Lock all EDR related cells & Unlock all UJB related cells.
    If ActiveWorkbook.Worksheets("System Selection").UJBBox.Value = strYes Then
        With ActiveWorkbook.Worksheets("Gas Analyzer")
            .Range("A16:A23").Locked = False
            .Range("D16:D23").Locked = False
            .Range("A27:A28").Locked = False
            .Range("D27:D28").Locked = False
        End With
    End If

    Call GasAnalyzerCopyData

Application.EnableEvents = True

End Sub
