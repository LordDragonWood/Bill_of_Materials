Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Option Explicit

    Dim ESRFlowPN As String
    Dim ESRUPSPN As String
    Dim ESRSideKickPN As String
    Dim ESRTorqueEPN As String
'   Laptop Computer Running Windows 7 (ViewSonic) (Inside Argentina)    CMP229

Private Sub Worksheet_Activate()
'Created for Pason by Dragon Wood (October 2015)
'Automatically sets the Required equipment to "Yes" and sets the default value for each. These can of course be changed.

Application.EnableEvents = False

    Call ShowLegend

    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt1 Then
        ESRFlowPN = "FLOW008"
        ESRUPSPN = "PWRASS099"
        ESRSideKickPN = "DHC029"
        ESRTorqueEPN = "SEN008"
    ElseIf ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt2 Then
        ESRFlowPN = "FLOW009"
        ESRUPSPN = "PWR095"
        ESRSideKickPN = "DHC033"
        ESRTorqueEPN = "SENASS109"
    End If

    'Only choose the 110V items if 110V is selected.
    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt1 Then
        With ActiveWorkbook.Worksheets("ESR")
            .Range("C8").Value = ESRFlowPN
            .Range("C22").Value = ESRTorqueEPN
            .Range("C24").Value = ESRSideKickPN
            .Range("C27").Value = ESRUPSPN
            .Range("C115").Value = ADRPowerBarPN1
        End With
    End If

    'Only choose 220V if 220V is selected.
    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt2 Then
        With ActiveWorkbook.Worksheets("ESR")
            .Range("C8").Value = ESRFlowPN
            .Range("C22").Value = ESRTorqueEPN
            .Range("C24").Value = ESRSideKickPN
            .Range("C27").Value = ESRUPSPN
            .Range("C115").Value = ADRPowerBarPN2
        End With
    End If

    'Warn the user they cannot use an ESR without a UJB.
    If ActiveWorkbook.Worksheets("System Selection").UJBBox.Value = strNo Then ESRUJB

    Call ESRCopyData

Application.EnableEvents = True

End Sub
