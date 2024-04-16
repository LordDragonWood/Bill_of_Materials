Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Option Explicit

    Dim PRDTypePN As String

Private Sub Worksheet_Activate()
'Created for Pason by Dragon Wood (August 2015)
'Automatically sets the Required equipment to "Yes" and sets the default value for each. These can of course be changed.

Application.EnableEvents = False

    Call ShowLegend

    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt1 Then
        PRDTypePN = "PRD001"
    ElseIf ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt2 Then
        PRDTypePN = "PRD002"
    End If

    'Only choose the 110V items if 110V is selected.
    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt1 Then
        With ActiveWorkbook.Worksheets("PRD")
            .Range("C3").Value = PRDTypePN
            .Range("C22").Value = ADRPowerBarPN1
        End With
    End If

    'Only choose 220V if 220V is selected.
    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt2 Then
        With ActiveWorkbook.Worksheets("PRD")
            .Range("C3").Value = PRDTypePN
            .Range("C22").Value = ADRPowerBarPN2
        End With
    End If

    'Warn the user they cannot use an PRD without a UJB.
    If ActiveWorkbook.Worksheets("System Selection").UJBBox.Value = strNo Then PRDUJB

    Call PRDCopyData

Application.EnableEvents = True

End Sub
