Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Option Explicit

Private Sub Worksheet_Activate()
'Created for Pason by Dragon Wood (August 2015)
'Automatically sets the Required equipment to "Yes" and sets the default value for each. These can of course be changed.

Application.EnableEvents = False

    Call ShowLegend

    'Only choose the 110V items if 110V is selected.
    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt1 Then
        With ActiveWorkbook.Worksheets("Workstations")
            .Range("C20").Value = CMPPowerCordPN1
            .Range("C25").Value = SurgeSupressorPN1
        End With
    End If

    'Only choose the 220V items if 220V is selected.
    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt2 Then
        With ActiveWorkbook.Worksheets("Workstations")
            .Range("C20").Value = CMPPowerCordPN2
            .Range("C25").Value = SurgeSupressorPN2
        End With
    End If

    Call WorkstationsCopyData

Application.EnableEvents = True

End Sub
