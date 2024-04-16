Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Option Explicit

    Dim ADRConBoxPN As String
    Dim ADRSideKickPN As String

Private Sub Worksheet_Activate()
'Created for Pason by Dragon Wood (August 2015)
'Automatically selects the required equipment based on voltage.
'Locks or unlocks equipment based on voltage.

Application.EnableEvents = False

    Call ShowLegend
    
    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt1 Then
        ADRConBoxPN = "ADR030"
        ADRSideKickPN = "DHC029"
    ElseIf ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt2 Then
        ADRConBoxPN = "ADR034"
        ADRSideKickPN = "DHC033"
    End If
            
    'Only choose the 110V items if 110V is selected, choose 220V if 220V is selected.
    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt1 Then
        With ActiveWorkbook.Worksheets("AutoDriller")
            .Range("C3").Value = ADRConBoxPN
            .Range("C10").Value = ADRSideKickPN
            .Range("C48").Value = ADRPowerBarPN1
        End With
    End If

    'Lock and clear the 110V quantity cells and unlock and set the 220V cells.
    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt2 Then
        With ActiveWorkbook.Worksheets("AutoDriller")
            .Range("C3").Value = ADRConBoxPN
            .Range("C10").Value = ADRSideKickPN
            .Range("C48").Value = ADRPowerBarPN2
        End With
    End If
    
    'Lock the SideKick cells if a SideKick is not selected.
    If ActiveWorkbook.Worksheets("System Selection").SideKickBox.Value = strNo Then
        With ActiveWorkbook.Worksheets("AutoDriller")
            .Range("A10").Value = strNo
            .Range("A10").Locked = True
            .Range("D10").Locked = True
            .Range("D10").ClearContents
        End With
    End If
    
    'Unlock the SideKick cells if the SideKick is selected.
    If ActiveWorkbook.Worksheets("System Selection").SideKickBox.Value = strYes Then
        With ActiveWorkbook.Worksheets("AutoDriller")
            .Range("A10").Value = strYes
            .Range("A10").Locked = False
            .Range("D10").Locked = False
        End With
    End If
    
    'Lock all UJB related cells & Unlock all EDR related cells.
    If ActiveWorkbook.Worksheets("System Selection").UJBBox.Value = strNo Then
        With ActiveWorkbook.Worksheets("AutoDriller")
            .Range("A40:A47").Locked = True
            .Range("A40:A47").Value = strNo
            .Range("D40:D47").ClearContents
            .Range("D40:D47").Locked = True
            .Range("A49:A50").Locked = True
            .Range("A49:A50").Value = strNo
            .Range("D49:D50").Locked = True
            .Range("D49:D50").ClearContents
        End With
    End If

    'Lock all EDR related cells & Unlock all UJB related cells.
    If ActiveWorkbook.Worksheets("System Selection").UJBBox.Value = strYes Then
        With ActiveWorkbook.Worksheets("AutoDriller")
            .Range("A40:A47").Locked = False
            .Range("D40:D47").Locked = False
            .Range("A49:A50").Locked = False
            .Range("D49:D50").Locked = False
        End With
    End If
    
    Call AutoDrillerCopyData
    
    Application.EnableEvents = True

End Sub
