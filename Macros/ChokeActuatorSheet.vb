Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Option Explicit

    Dim CAAssemblyPN As String

Private Sub Worksheet_Activate()
'Created for Pason by Dragon Wood (August 2015)
'Automatically sets the Required equipment to "Yes" and sets the default value for each. These can of course be changed.

Application.EnableEvents = False
    
    Call ShowLegend
    
    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt1 Then
        CAAssemblyPN = "SUBASS018"
    ElseIf ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt2 Then
        CAAssemblyPN = "SUBASS123"
    End If

    'Only choose the 110V items if 110V is selected.
    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt1 Then
        With ActiveWorkbook.Worksheets("Choke Actuator")
            .Range("C2").Value = CAAssemblyPN
            .Range("C40").Value = ADRPowerBarPN1
        End With
    End If
    
    'Only choose the 220V items if 220V is selected.
    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt2 Then
        With ActiveWorkbook.Worksheets("Choke Actuator")
            .Range("C2").Value = CAAssemblyPN
            .Range("C40").Value = ADRPowerBarPN2
        End With
    End If
    
    'Lock all UJB related cells & Unlock all EDR related cells.
    If ActiveWorkbook.Worksheets("System Selection").UJBBox.Value = strNo Then
        With ActiveWorkbook.Worksheets("Choke Actuator")
            .Range("A30:A37").Locked = True
            .Range("A30:A37").Value = strNo
            .Range("D30:D37").Locked = True
            .Range("D30:D37").ClearContents
            .Range("A41:A42").Locked = True
            .Range("A41:A42").Value = strNo
            .Range("D41:D42").Locked = True
            .Range("D41:D42").ClearContents
        End With
    End If

    'Lock all EDR related cells & Unlock all UJB related cells.
    If ActiveWorkbook.Worksheets("System Selection").UJBBox.Value = strYes Then
        With ActiveWorkbook.Worksheets("Choke Actuator")
            .Range("A30:A37").Locked = False
            .Range("D30:D37").Locked = False
            .Range("A41:A42").Locked = False
            .Range("D41:D42").Locked = False
        End With
    End If

    Call ChokeCopyData

Application.EnableEvents = True

End Sub
