Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Option Explicit

    Dim PVTMonitorPN As String
    Dim PVTFlowPN As String

Private Sub Worksheet_Activate()
'Created for Pason by Dragon Wood (August 2015)
'Automatically sets the Required equipment to "Yes" and sets the default value for each. These can of course be changed.

Application.EnableEvents = False

    Call ShowLegend

    If ActiveWorkbook.Worksheets("System Selection").UnitsBox.Value = strUnitsI Then Call PVTImperial
    If ActiveWorkbook.Worksheets("System Selection").UnitsBox.Value = strUnitsM Then Call PVTMetric
    If ActiveWorkbook.Worksheets("System Selection").UJBBox.Value = strYes Then Call PVTUJB
    If ActiveWorkbook.Worksheets("System Selection").UJBBox.Value = strNo Then Call PVTNoUJB

    'Only choose the 110V items if 110V is selected.
    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt1 Then
        With ActiveWorkbook.Worksheets("PVT")
            .Range("C6").Value = PVTMonitorPN
            .Range("C7").Value = PVTFlowPN
            .Range("C52").Value = ADRPowerBarPN1
        End With
    End If

    'Only choose 220V if 220V is selected.
    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt2 Then
        With ActiveWorkbook.Worksheets("PVT")
            .Range("C6").Value = PVTMonitorPN
            .Range("C7").Value = PVTFlowPN
            .Range("C52").Value = ADRPowerBarPN2
        End With
    End If

    'Lock all UJB related cells & Unlock all EDR related cells.
    If ActiveWorkbook.Worksheets("System Selection").UJBBox.Value = strNo Then
        With ActiveWorkbook.Worksheets("PVT")
            .Range("A10:A16").Locked = False
            .Range("D10:D16").Locked = False
            .Range("A18:A19").Locked = False
            .Range("D18:D19").Locked = False
            .Range("A27").Locked = True
            .Range("A27").Value = strNo
            .Range("D27").Locked = True
            .Range("D27").ClearContents
            .Range("A41:A48").Locked = True
            .Range("A41:A48").Value = strNo
            .Range("D41:D48").Locked = True
            .Range("D41:D48").ClearContents
            .Range("A58").Locked = True
            .Range("A58").Value = strNo
            .Range("D58").Locked = True
            .Range("D58").ClearContents
        End With
    End If

    'Lock all EDR related cells & Unlock all UJB related cells.
    If ActiveWorkbook.Worksheets("System Selection").UJBBox.Value = strYes Then
        With ActiveWorkbook.Worksheets("PVT")
            .Range("A10:A16").Locked = True
            .Range("A10:A16").Value = strNo
            .Range("D10:D16").Locked = True
            .Range("D10:D16").ClearContents
            .Range("A18:A19").Locked = True
            .Range("A18:A19").Value = strNo
            .Range("D18:D19").Locked = True
            .Range("D18:D19").ClearContents
            .Range("A27").Locked = False
            .Range("D27").Locked = False
            .Range("A41:A48").Locked = False
            .Range("D41:D48").Locked = False
            .Range("A58").Locked = False
            .Range("D58").Locked = False
        End With
    End If

    'Lock all Mud Probe related cells & Unlock all Radar Probe related cells.
    If ActiveWorkbook.Worksheets("System Selection").ProbeBox.Value = strPrbR Then
        With ActiveWorkbook.Worksheets("PVT")
            .Range("A8").Locked = True
            .Range("A8").Value = strNo
            .Range("D8").Locked = True
            .Range("D8").ClearContents
            .Range("A9").Locked = False
            .Range("D9").Locked = False
            .Range("A17").Locked = True
            .Range("A17").Value = strNo
            .Range("D17").Locked = True
            .Range("D17").ClearContents
            .Range("A20:A24").Locked = True
            .Range("A20:A24").Value = strNo
            .Range("D20:D24").Locked = True
            .Range("D20:D24").ClearContents
            .Range("A25").Locked = False
            .Range("D25").Locked = False
            .Range("A26").Locked = True
            .Range("A26").Value = strNo
            .Range("D26").Locked = True
            .Range("D26").ClearContents
            .Range("A28").Locked = False
            .Range("A28").Value = strYes
            .Range("D28").Locked = False
            .Range("D28").Value = 1
        End With
    End If

    'Lock all Radar Probe related cells & Unlock all Mud Probe related cells.
    If ActiveWorkbook.Worksheets("System Selection").ProbeBox.Value = strPrbM Then
        With ActiveWorkbook.Worksheets("PVT")
            .Range("A8").Locked = False
            .Range("D8").Locked = False
            .Range("A9").Locked = True
            .Range("A9").Value = strNo
            .Range("D9").Locked = True
            .Range("D9").ClearContents
            .Range("A17").Locked = False
            .Range("A17").Value = strYes
            .Range("D17").Locked = False
            .Range("D17").Value = 1
            .Range("A20:A24").Locked = False
            .Range("D20:D24").Locked = False
            .Range("A25").Locked = True
            .Range("A25").Value = strNo
            .Range("D25").Locked = True
            .Range("D25").ClearContents
            .Range("A26").Locked = False
            .Range("D26").Locked = False
            .Range("A28").Locked = True
            .Range("A28").Value = strNo
            .Range("D28").Locked = True
            .Range("D28").ClearContents
        End With
    End If

    'Unlock all Probe related cells.
    If ActiveWorkbook.Worksheets("System Selection").ProbeBox.Value = strPrbB Then
        With ActiveWorkbook.Worksheets("PVT")
            .Range("A8:A9").Locked = False
            .Range("D8:D9").Locked = False
            .Range("A17").Locked = False
            .Range("A17").Value = strYes
            .Range("D17").Locked = False
            .Range("D17").Value = 1
            .Range("A20:A26").Locked = False
            .Range("D20:D26").Locked = False
            .Range("A28").Locked = False
            .Range("A28").Value = strYes
            .Range("D28").Locked = False
            .Range("D28").Value = 1
        End With
    End If

    'Select the Flow parts if Flow is selelcted
    If ActiveWorkbook.Worksheets("System Selection").FlowBox.Value = strYes Then
        With ActiveWorkbook.Worksheets("PVT")
            .Range("A3:A5").Value = strYes
            .Range("A3:A5").Locked = False
            .Range("D3:D5").Locked = False
            .Range("A7").Locked = False
            .Range("A7").Value = strYes
            .Range("D7").Locked = False
        End With
    End If

    'Unselect the Flow parts if Flow is not selelcted
    If ActiveWorkbook.Worksheets("System Selection").FlowBox.Value = strNo Then
        With ActiveWorkbook.Worksheets("PVT")
            .Range("A3:A5").Value = strNo
            .Range("A3:A5").Locked = True
            .Range("D3:D5").Locked = True
            .Range("D3:D5").ClearContents
            .Range("A7").Locked = True
            .Range("A7").Value = strNo
            .Range("D7").Locked = True
            .Range("D7").ClearContents
        End With
    End If

    Call PVTCopyData

Application.EnableEvents = True

End Sub
Function PVTImperial()

    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt1 Then
        PVTMonitorPN = "PVTASS008"
    ElseIf ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt2 Then
        PVTMonitorPN = "PVTASS009"
    End If

End Function

Function PVTMetric()

    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt1 Then
        PVTMonitorPN = "PVTASS005"
    ElseIf ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt2 Then
        PVTMonitorPN = "PVTASS006"
    End If

End Function

Function PVTUJB()

    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt1 Then
        PVTFlowPN = "FLOW008"
    ElseIf ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt2 Then
        PVTFlowPN = "FLOW009"
    End If

End Function

Function PVTNoUJB()

    PVTFlowPN = "FLOW004"

End Function
