Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Option Explicit

Dim EDRDHCPN As String
Dim EDRUPSPN As String
Dim EDRJBoxPN As String
Dim EDRSubJBoxPN As String

Private Sub Worksheet_Activate()
'Created for Pason by Dragon Wood (August 2015)
'Automatically sets the Required equipment to "Yes" and sets the default value for each. These can of course be changed.

Application.EnableEvents = False

    Call ShowLegend

    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt1 Then
        EDRDHCPN = "DHC030"
        EDRUPSPN = "PWRASS099"
        EDRJBoxPN = "DHC010"
        EDRSubJBoxPN = "DHC005"
    ElseIf ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt2 Then
        EDRDHCPN = "DHC034"
        EDRUPSPN = "PWR095"
        EDRJBoxPN = "DHC013"
        EDRSubJBoxPN = "DHC014"
    End If

    'Only choose the 110V items if 110V is selected.
    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt1 Then
        With ActiveWorkbook.Worksheets("EDR")
            .Range("C9").Value = EDRDHCPN
            .Range("C10").Value = EDRJBoxPN
            .Range("C11").Value = EDRSubJBoxPN
            .Range("C20").Value = TorqueEPN1
            .Range("C26").Value = EDRUPSPN
            .Range("C89").Value = SurgeSupressorPN1
            .Range("C90").Value = ADRPowerBarPN1
            .Range("C115").Value = CMPPowerCordPN1
        End With
    End If

    'Only choose 220V if 220V is selected.
    If ActiveWorkbook.Worksheets("System Selection").VoltageBox.Value = strVolt2 Then
        With ActiveWorkbook.Worksheets("EDR")
            .Range("C9").Value = EDRDHCPN
            .Range("C10").Value = EDRJBoxPN
            .Range("C11").Value = EDRSubJBoxPN
            .Range("C20").Value = TorqueEPN2
            .Range("C26").Value = EDRUPSPN
            .Range("C89").Value = SurgeSupressorPN2
            .Range("C90").Value = ADRPowerBarPN2
            .Range("C115").Value = CMPPowerCordPN2
        End With
    End If

    'Lock all UJB related cells & Unlock all EDR related cells.
    If ActiveWorkbook.Worksheets("System Selection").UJBBox.Value = strNo Then
        With ActiveWorkbook.Worksheets("EDR")
            .Range("A3").Locked = False
            .Range("A3").Value = strYes
            .Range("D3").Locked = False
            .Range("D3").Value = 1
            .Range("A10:A11").Locked = False
            .Range("A10:A11").Value = strYes
            .Range("D10:D11").Locked = False
            .Range("D10:D11").Value = 1
            .Range("A23:A24").Locked = True
            .Range("A23:A24").Value = strNo
            .Range("D23:D24").Locked = True
            .Range("D23:D24").ClearContents
            .Range("A62").Locked = False
            .Range("D62").Locked = False
            .Range("A69:A76").Value = strNo
            .Range("A69:A76").Locked = True
            .Range("D69:D76").Locked = True
            .Range("D69:D76").ClearContents
        End With
    End If

    'Lock all EDR related cells & Unlock all UJB related cells.
    If ActiveWorkbook.Worksheets("System Selection").UJBBox.Value = strYes Then
        With ActiveWorkbook.Worksheets("EDR")
            .Range("A3").Locked = True
            .Range("A3").Value = strNo
            .Range("D3").Locked = True
            .Range("D3").ClearContents
            .Range("A10:A11").Locked = True
            .Range("A10:A11").Value = strNo
            .Range("D10:D11").Locked = True
            .Range("D10:D11").ClearContents
            .Range("A23:A24").Locked = False
            .Range("A23:A24").Value = strYes
            .Range("D23:D24").Locked = False
            .Range("A62").Locked = True
            .Range("A62").Value = strNo
            .Range("D62").Locked = True
            .Range("D62").ClearContents
            .Range("A69:A76").Locked = False
            .Range("D69:D76").Locked = False
        End With
    End If

    Call EDRCopyData

Application.EnableEvents = True

End Sub
