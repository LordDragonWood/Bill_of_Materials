Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Option Explicit

    Dim ePVTFlowPN As String
    Dim ePVTPRDPN As String

Private Sub Worksheet_Activate()
'Created for Pason by Dragon Wood (August 2015)
'Automatically sets the Required equipment to "Yes" and sets the default value for each. These can of course be changed.

Application.EnableEvents = False

    Call ShowLegend
    
    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt1 Then
        ePVTFlowPN = "FLOW008"
        ePVTPRDPN = "PRD001"
    ElseIf ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt2 Then
        ePVTFlowPN = "FLOW009"
        ePVTPRDPN = "PRD002"
    End If

    'Only choose the 110V items if 110V is selected.
    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt1 Then
        With ActiveWorkbook.Worksheets("ePVT")
            .Range("C6").Value = ePVTPRDPN
            .Range("C7").Value = ePVTFlowPN
            .Range("C49").Value = ADRPowerBarPN1
        End With
    End If

    'Only choose 220V if 220V is selected.
    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt2 Then
        With ActiveWorkbook.Worksheets("ePVT")
            .Range("C6").Value = ePVTPRDPN
            .Range("C7").Value = ePVTFlowPN
            .Range("C49").Value = ADRPowerBarPN2
        End With
    End If

    'Warn the user they cannot use an ePVT without a UJB.
    If ActiveWorkbook.Worksheets("System Selection").UJBBox.Value = strNo Then ePVTUJB

    'Lock all Mud Probe related cells & Unlock all Radar Probe related cells.
    If ActiveWorkbook.Worksheets("System Selection").ProbeBox.Value = strPrbR Then
        With ActiveWorkbook.Worksheets("ePVT")
            .Range("A9").Locked = True
            .Range("A9").Value = strNo
            .Range("D9").Locked = True
            .Range("D9").ClearContents
            .Range("A10").Locked = False
            .Range("D10").Locked = False
            .Range("A11:A16").Locked = True
            .Range("A11:A16").Value = strNo
            .Range("D11:D16").Locked = True
            .Range("D11:D16").ClearContents
            .Range("A17").Locked = False
            .Range("D17").Locked = False
            .Range("A18").Locked = True
            .Range("A18").Value = strNo
            .Range("D18").Locked = True
            .Range("D18").ClearContents
            .Range("A22").Locked = False
            .Range("A22").Value = strYes
            .Range("D22").Locked = False
            .Range("D22").Value = 1
        End With
    End If

    'Lock all Radar Probe related cells & Unlock all Mud Probe related cells.
    If ActiveWorkbook.Worksheets("System Selection").ProbeBox.Value = strPrbM Then
        With ActiveWorkbook.Worksheets("ePVT")
            .Range("A9").Locked = False
            .Range("D9").Locked = False
            .Range("A10").Locked = True
            .Range("A10").Value = strNo
            .Range("D10").Locked = True
            .Range("D10").ClearContents
            .Range("A11:A16").Locked = False
            .Range("A11").Value = strYes
            .Range("D11:D16").Locked = False
            .Range("D11").Value = 1
            .Range("A17").Locked = True
            .Range("A17").Value = strNo
            .Range("D17").Locked = True
            .Range("D17").ClearContents
            .Range("A18").Locked = False
            .Range("D18").Locked = False
            .Range("A22").Locked = True
            .Range("A22").Value = strNo
            .Range("D22").Locked = True
            .Range("D22").ClearContents
        End With
    End If

    'Unlock all Probe related cells.
    If ActiveWorkbook.Worksheets("System Selection").ProbeBox.Value = strPrbB Then
        With ActiveWorkbook.Worksheets("ePVT")
            .Range("A9").Locked = False
            .Range("D9").Locked = False
            .Range("A10").Locked = False
            .Range("D10").Locked = False
            .Range("A11:A16").Locked = False
            .Range("A11").Value = strYes
            .Range("D11:D16").Locked = False
            .Range("D11").Value = 1
            .Range("A17").Locked = False
            .Range("D17").Locked = False
            .Range("A18").Locked = False
            .Range("D18").Locked = False
            .Range("A22").Locked = False
            .Range("A22").Value = strYes
            .Range("D22").Locked = False
            .Range("D22").Value = 1
        End With
    End If
    
    'Select the Flow parts if Flow is selelcted
    If ActiveWorkbook.Worksheets("System Selection").FlowBox.Value = strYes Then
        With ActiveWorkbook.Worksheets("ePVT")
            .Range("A2:A4").Value = strYes
            .Range("A2:A4").Locked = False
            .Range("D2:D4").Locked = False
            .Range("A7").Locked = False
            .Range("A7").Value = strYes
            .Range("D7").Locked = False
        End With
    End If
    
    'Unselect the Flow parts if Flow is not selelcted
    If ActiveWorkbook.Worksheets("System Selection").FlowBox.Value = strNo Then
        With ActiveWorkbook.Worksheets("ePVT")
            .Range("A2:A4").Value = strNo
            .Range("A2:A4").Locked = True
            .Range("D2:D4").Locked = True
            .Range("D2:D4").ClearContents
            .Range("A7").Locked = True
            .Range("A7").Value = strNo
            .Range("D7").Locked = True
            .Range("D7").ClearContents
        End With
    End If

    Call ePVTCopyData

Application.EnableEvents = True

End Sub
