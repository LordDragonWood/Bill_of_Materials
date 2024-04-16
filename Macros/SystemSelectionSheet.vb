Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1

Private Sub AutoDrillerBox_Change()
'Created for Pason by Dragon Wood (August 2015)
'Hides or Unhides the right sheet

    If LCase(AutoDrillerBox.Value) = "yes" And Worksheets("AutoDriller").Visible = False Then
        Worksheets("AutoDriller").Visible = True
    End If
    If LCase(AutoDrillerBox.Value) = "yes" Then AutoDrillerUnhide
    If LCase(AutoDrillerBox.Value) = "no" And Worksheets("AutoDriller").Visible = True Then
        Call AutoDrillerClear
        Worksheets("AutoDriller").Visible = False
    End If
End Sub

Private Sub CasingBox_Change()
'Created for Pason by Dragon Wood (July 2015)
'Hides or Unhides the right sheet

    If LCase(CasingBox.Value) = "yes" And Worksheets("Casing Pressure").Visible = False Then
        Worksheets("Casing Pressure").Visible = True
    End If
    If LCase(CasingBox.Value) = "yes" Then CasingPressureUnhide
    If LCase(CasingBox.Value) = "no" And Worksheets("Casing Pressure").Visible = True Then
        Call CasingPressureClear
        Worksheets("Casing Pressure").Visible = False
    End If
End Sub

Private Sub ChokeBox_Change()
'Created for Pason by Dragon Wood (July 2015)
'Hides or Unhides the right sheet

    If LCase(ChokeBox.Value) = "yes" And Worksheets("Choke Actuator").Visible = False Then
        Worksheets("Choke Actuator").Visible = True
    End If
    If LCase(ChokeBox.Value) = "yes" Then ChokeActuatorUnhide
    If LCase(ChokeBox.Value) = "no" And Worksheets("Choke Actuator").Visible = True Then
        Call ChokeActuatorClear
        Worksheets("Choke Actuator").Visible = False
    End If
End Sub

Private Sub EDRBox_Change()
'Created for Pason by Dragon Wood (July 2015)
'Hides or Unhides the right sheet

    If LCase(EDRBox.Value) = "yes" And Worksheets("EDR").Visible = False Then
        Worksheets("EDR").Visible = True
    End If
    If LCase(EDRBox.Value) = "yes" Then EDRUnhide
    If LCase(EDRBox.Value) = "no" And Worksheets("EDR").Visible = True Then
        Call EDRClear
        Worksheets("EDR").Visible = False
    End If
End Sub

Private Sub ePVTBox_Change()
'Created for Pason by Dragon Wood (July 2015)
'Hides or Unhides the right sheet

    If LCase(ePVTBox.Value) = "yes" And Worksheets("ePVT").Visible = False Then
        Worksheets("ePVT").Visible = True
    End If
    If LCase(ePVTBox.Value) = "yes" Then ePVTUnhide
    If LCase(ePVTBox.Value) = "no" And Worksheets("ePVT").Visible = True Then
        Call ePVTClear
        Worksheets("ePVT").Visible = False
    End If
End Sub

Private Sub ESRBox_Change()
'Created for Pason by Dragon Wood (September 2015)
'Hides or Unhides the right sheet

    If LCase(ESRBox.Value) = "yes" And Worksheets("ESR").Visible = False Then
        Worksheets("ESR").Visible = True
    End If
    If LCase(ESRBox.Value) = "yes" Then ESRUnhide
    If LCase(ESRBox.Value) = "no" And Worksheets("ESR").Visible = True Then
        Call ESRClear
        Worksheets("ESR").Visible = False
    End If
End Sub

Private Sub GABox_Change()
'Created for Pason by Dragon Wood (July 2015)
'Hides or Unhides the right sheet

    If LCase(GABox.Value) = "yes" And Worksheets("Gas Analyzer").Visible = False Then
        Worksheets("Gas Analyzer").Visible = True
    End If
    If LCase(GABox.Value) = "yes" Then GasAnalyzerUnhide
    If LCase(GABox.Value) = "no" And Worksheets("Gas Analyzer").Visible = True Then
        Call GasAnalyzerClear
        Worksheets("Gas Analyzer").Visible = False
    End If
End Sub

Private Sub HGasBox_Change()
'Created for Pason by Dragon Wood (July 2015)
'Hides or Unhides the right sheet

    If LCase(HGasBox.Value) = "yes" And Worksheets("H-Gas").Visible = False Then
        Worksheets("H-Gas").Visible = True
    End If
    If LCase(HGasBox.Value) = "yes" Then HGasUnhide
    If LCase(HGasBox.Value) = "no" And Worksheets("H-Gas").Visible = True Then
        Call HGasClear
        Worksheets("H-Gas").Visible = False
    End If
End Sub

Private Sub PRDBox_Change()
'Created for Pason by Dragon Wood (July 2015)
'Hides or Unhides the right sheet

    If LCase(PRDBox.Value) = "yes" And Worksheets("PRD").Visible = False Then
        Worksheets("PRD").Visible = True
    End If
    If LCase(PRDBox.Value) = "yes" Then PRDUnhide
    If LCase(PRDBox.Value) = "no" And Worksheets("PRD").Visible = True Then
        Call PRDClear
        Worksheets("PRD").Visible = False
    End If
End Sub

Private Sub PVTBox_Change()
'Created for Pason by Dragon Wood (July 2015)
'Hides or Unhides the right sheet

    If LCase(PVTBox.Value) = "yes" And Worksheets("PVT").Visible = False Then
        Worksheets("PVT").Visible = True
    End If
    If LCase(PVTBox.Value) = "yes" Then PVTUnhide
    If LCase(PVTBox.Value) = "no" And Worksheets("PVT").Visible = True Then
        Call PVTClear
        Worksheets("PVT").Visible = False
    End If
End Sub

Private Sub WorkstationsBox_Change()
'Created for Pason by Dragon Wood (July 2015)
'Hides or Unhides the right sheet

    If LCase(WorkstationsBox.Value) = "yes" And Worksheets("Workstations").Visible = False Then
        Worksheets("Workstations").Visible = True
    End If
    If LCase(WorkstationsBox.Value) = "yes" Then WorkstationsUnhide
    If LCase(WorkstationsBox.Value) = "no" And Worksheets("Workstations").Visible = True Then
        Call WorkstationsClear
        Worksheets("Workstations").Visible = False
    End If
End Sub

Private Sub btnSaveProject_Click()
'Created for Pason by Dragon Wood (October 2015).
    Dim GetBook As String
    Dim iDot As Long
    
    GetBook = ThisWorkbook.Name

    iDot = InStrRev(GetBook, ".")
    
    GetBook = Left(GetBook, iDot - 1)

    If GetBook = "Pason Installation BOMs" Then
        Call SaveProjectAs
    Else
        Call SaveWorkbook
    End If

End Sub
