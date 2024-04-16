Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Public Sub EDRReset(control As IRibbonControl)
'Created for Pason by Dragon Wood (August 2015)
'Resets the EDR page so it can be used again.

    Application.ScreenUpdating = False

    'Clear all data on the page.
    Call EDRClear
'    Call EDRCopyData

    'Close the sheet.
    With ActiveWorkbook.Worksheets("System Selection")
        .EDRBox.Value = strNo
    End With

    'Reopen the sheet to refill the required parts.
    With ActiveWorkbook.Worksheets("System Selection")
        .EDRBox.Value = strYes
    End With

    Application.ScreenUpdating = True
    Sheets("System Selection").Select

End Sub

Function EDRUnhide()
'Created for Pason by Dragon Wood (August 2015)
'Automatically sets the Required equipment to "Yes" and sets the default value for each. These can of course be changed.

    'Ensure the EDR sheet is selected.
    With ActiveWorkbook.Worksheets("EDR")

        'Set the Required equipment to Yes.
        .Range("A2:A27").Value = strYes

        'Set the default quantity for each item on the list
        .Range("D2:D3").Value = 1
        .Range("D4").Value = 2
        .Range("D5:D11").Value = 1
        .Range("D12").Value = 2
        .Range("D13:D27").Value = 1
    End With

End Function

Function EDRClear()
'Created for Pason by Dragon Wood (August 2015)
'Resets the EDR Page to the original unused state.

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    'Reset the Yes/No field to No on the EDR Sheet and clear the Quantity Needed Column
    With ActiveWorkbook.Worksheets("EDR")
        .Visible = True
        .Range("A2:A120").Value = strNo
        .Range("D2:D120").ClearContents
    End With

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    Application.GoTo Sheets("EDR").Range("A1"), True
    Application.GoTo Sheets("System Selection").Range("A1"), True

End Function

Public Sub EDRCheck()
'Created for Pason by Dragon Wood (August 2015)
'Checks to make sure that the items that are required, but need a choice are chosen.

    Call EDRDepthBracketABCheck
    Call EDRDepthBracketCCheck
    Call EDRDepthBracketDCheck
    Call EDRDepthBracketECheck
    Call RotaryBracketCheck
    Call DepthCableCheck
    Call RotaryTorqueCableCheck
    Call EDRTPCMonitorCheck
    Call EDRWSMonitorCheck
    Call DepthSensorCheck
    Call RotarySensorCheck
    Call TorqueSensorCheck
    Call EDRKeyboardCheck
    Call EDRMouseCheck
    Call EDRMousePadCheck
    Call EDRPowerCordCheck
    Call EDRSurgeSupressorCheck
    Call DHCWallMountCheck
    Call CirronetRadioCheck
    Call EDRSubUJBPigtailCheck

End Sub

Private Sub CirronetRadioCheck()
   
    If PreviewOrderRunControl = False Then
        Call CirronetRadioIncrement
        Call SetPreviewOrderRun
    Else
        Exit Sub
    End If

End Sub

Function SetPreviewOrderRun()

        PreviewOrderRunControl = ActiveWorkbook.Worksheets("Rig Survey Form").Range("AG14") = True
        
End Function

Function CirronetRadioIncrement()

    Dim CommCableAdd As Integer
    Dim CirronetRadioAdd As Integer

    CommCableAdd = ActiveWorkbook.Worksheets("EDR").Range("D56").Value
    CirronetRadioAdd = ActiveWorkbook.Worksheets("EDR").Range("D8").Value
    
    If ActiveWorkbook.Worksheets("EDR").Range("A8").Value = strYes Then
        With ActiveWorkbook.Worksheets("EDR")
            .Range("A56").Value = strYes
            .Range("D56").Value = CommCableAdd + 1
            .Range("D8").Value = CirronetRadioAdd + 1
        End With
    End If

End Function

Private Sub RotaryBracketCheck()

    'Check for a Rotary Bracket to be ordered.
    With ActiveWorkbook.Worksheets("EDR")
        For Each cell In Range("A30:A31")
            If cell = strYes Then Exit Sub
                Next cell
                Call RotaryBracketMessage
    End With

End Sub

Function RotaryBracketMessage()
'Created for Pason by Dragon Wood (October 2015)
'Warns the user they have not ordered a Rotary Bracket.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Rotary Bracket."
    Title = "EDR"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("EDR").Select

End Function

Private Sub DepthCableCheck()

    'Check for a Depth Cable to be ordered.
    With ActiveWorkbook.Worksheets("EDR")
        For Each cell In Range("A32:A34")
            If cell = strYes Then Exit Sub
                Next cell
                Call DepthCableMessage
    End With

End Sub

Function DepthCableMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a Depth Cable.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Depth Cable."
    Title = "EDR"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("EDR").Select

End Function

Private Sub RotaryTorqueCableCheck()

    'Check for a Rotary & Torque Cable to be ordered.
    With ActiveWorkbook.Worksheets("EDR")
        For Each cell In Range("A35:A36")
            If cell = strYes Then Exit Sub
                Next cell
                Call RotaryTorqueCableMessage
    End With

End Sub

Function RotaryTorqueCableMessage()
'Created for Pason by Dragon Wood (October 2015)
'Warns the user they have not ordered a Rotary & Torque Cable.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Rotary & Torque Cable."
    Title = "EDR"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("EDR").Select

End Function

Private Sub DepthSensorCheck()

    'Check for a Depth Sensor to be ordered.
    With ActiveWorkbook.Worksheets("EDR")
        For Each cell In Range("A41:A46")
            If cell = strYes Then Exit Sub
                Next cell
                Call DepthSensorMessage
    End With

End Sub

Function DepthSensorMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a Depth Sensor.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Depth Sensor."
    Title = "EDR"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("EDR").Select

End Function

Private Sub RotarySensorCheck()

    'Check for a Rotary Sensor to be ordered.
    With ActiveWorkbook.Worksheets("EDR")
        For Each cell In Range("A47:A48")
            If cell = strYes Then Exit Sub
                Next cell
                Call RotarySensorMessage
    End With

End Sub

Function RotarySensorMessage()
'Created for Pason by Dragon Wood (October 2015)
'Warns the user they have not ordered a Rotary Sensor.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Rotary Sensor."
    Title = "EDR"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("EDR").Select

End Function

Private Sub TorqueSensorCheck()

    If ActiveWorkbook.Worksheets("EDR").Range("A20").Value = strYes Then Exit Sub

    'Check for a Torque Sensor to be ordered.
    With ActiveWorkbook.Worksheets("EDR")
        For Each cell In Range("A49:A50")
            If cell = strYes Then Exit Sub
                Next cell
                Call TorqueSensorMessage
    End With

End Sub

Function TorqueSensorMessage()
'Created for Pason by Dragon Wood (October 2015)
'Warns the user they have not ordered a Torque Sensor.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Torque Sensor."
    Title = "EDR"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("EDR").Select

End Function

Private Sub EDRTPCMonitorCheck()
'Created for Pason by Dragon Wood (August 2015)
'If a TPC is ordered, it checks to see if a monitor is ordered.

    If ActiveWorkbook.Worksheets("EDR").Range("A27").Value = strYes Then Exit Sub
    If ActiveWorkbook.Worksheets("EDR").Range("A25").Value = strNo Then Exit Sub

    'Check for a monitor to be ordered.
    With ActiveWorkbook.Worksheets("EDR")
        For Each cell In Range("A37:A40")
            If cell = strYes Then Exit Sub
                Next cell
                Call EDRTPCMonitorMessage
    End With
End Sub

Function EDRTPCMonitorMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user if they ordered a workstation but did not order a monitor

    'Declare the variables
    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExclBox As Integer

    'Set the Message Box settings.
    Msg = "You ordered a TPC, but did not order a monitor."
    Msg = Msg & vbNewLine & vbNewLine
    Msg = Msg & "Would you like to add a monitor to your order?"
    Title = "Warning!"
    Config = vbYesNo + vbExclamation + vbDefaultButton1
    ExclBox = MsgBox(Msg, Config, Title)

    'Change the monitor selection to Yes.
    If ExclBox = vbYes Then Call EDRTPCMonitorYes

    'Verify the user does not want a monitor.
    If ExclBox = vbNo Then Call EDRTPCMonitorNo

End Function

Function EDRTPCMonitorYes()
'Created for Pason by Dragon Wood (August 2015)
'Changes a montior selection to Yes

        With ActiveWorkbook.Worksheets("EDR")
            .Range("A40").Value = strYes
            .Range("D25").Copy
            .Range("D40").PasteSpecial xlPasteValues
        End With

End Function

Function EDRTPCMonitorNo()
'Created for Pason by Dragon Wood (August 2015)
'Verifies with the user they do not want a monitor.

    'Declare the variables
    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExclBox As Integer

    'Set the Message Box settings.
    Msg = "Are you sure you do not want a monitor?"
    Title = "Are You Sure?"
    Config = vbYesNo + vbQuestion + vbDefaultButton2
    ExclBox = MsgBox(Msg, Config, Title)

    'Change the monitor selection to Yes.
    If ExclBox = vbNo Then EDRMonitorYes

End Function

Private Sub EDRWSMonitorCheck()
'Created for Pason by Dragon Wood (August 2015)
'If a workstation is ordered, it checks to see if a monitor is ordered.

    If ActiveWorkbook.Worksheets("EDR").Range("A27").Value = strNo Then Exit Sub

    'Check for a monitor to be ordered.
    With ActiveWorkbook.Worksheets("EDR")
        For Each cell In Range("A37:A40")
            If cell = strYes Then Exit Sub
                Next cell
                Call EDRWSMonitorMessage
    End With
End Sub

Function EDRWSMonitorMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user if they ordered a workstation but did not order a monitor

    'Declare the variables
    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExclBox As Integer

    'Set the Message Box settings.
    Msg = "You ordered a workstation, but did not order a monitor."
    Msg = Msg & vbNewLine & vbNewLine
    Msg = Msg & "Would you like to add a monitor to your order?"
    Title = "Warning!"
    Config = vbYesNo + vbExclamation + vbDefaultButton1
    ExclBox = MsgBox(Msg, Config, Title)

    'Change the monitor selection to Yes.
    If ExclBox = vbYes Then Call EDRWSMonitorYes

    'Verify the user does not want a monitor.
    If ExclBox = vbNo Then Call EDRWSMonitorNo

End Function

Function EDRWSMonitorYes()
'Created for Pason by Dragon Wood (August 2015)
'Changes a montior selection to Yes

        With ActiveWorkbook.Worksheets("EDR")
            .Range("A40").Value = strYes
            .Range("D27").Copy
            .Range("D40").PasteSpecial xlPasteValues
        End With

End Function

Function EDRWSMonitorNo()
'Created for Pason by Dragon Wood (August 2015)
'Verifies with the user they do not want a monitor.

    'Declare the variables
    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExclBox As Integer

    'Set the Message Box settings.
    Msg = "Are you sure you do not want a monitor?"
    Title = "Are You Sure?"
    Config = vbYesNo + vbQuestion + vbDefaultButton2
    ExclBox = MsgBox(Msg, Config, Title)

    'Change the monitor selection to Yes.
    If ExclBox = vbNo Then EDRWSMonitorYes

End Function

Private Sub EDRKeyboardCheck()

    'Check to see if a Workstation is ordered.
    If ActiveWorkbook.Worksheets("EDR").Range("A27").Value = strNo Then Exit Sub

    'Check for a Keyboard to be ordered.
    With ActiveWorkbook.Worksheets("EDR")
        For Each cell In Range("A112")
            If cell = strNo Then Exit Sub
                Next cell
                Call EDRKeyboardMessage
    End With

End Sub

Function EDRKeyboardMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have ordered a Keyboard.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You do not need to order a keyboard when you order a Workstation."
    Msg = Msg & vbNewLine
    Msg = Msg & "The keyboard is part of the Workstation kit."
    Msg = Msg & vbNewLine & vbNewLine
    Msg = Msg & "Would you like to remove the keyboard from your order?"
    Title = "EDR"
    Config = vbYesNo + vbInformation + vbDefaultButton1
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("EDR").Select

    'Change the Keyboard selection to No.
    If ExcelBox = vbYes Then Call EDRKeyboardYes

End Function

Function EDRKeyboardYes()
'Created for Pason by Dragon Wood (August 2015)

    With ActiveWorkbook.Worksheets("EDR")
        .Range("A112").Value = strNo
        .Range("D112").ClearContents
    End With

End Function

Private Sub EDRMouseCheck()

    'Check to see if a Workstation is ordered.
    If ActiveWorkbook.Worksheets("EDR").Range("A27").Value = strNo Then Exit Sub

    'Check for a Mouse to be ordered.
    With ActiveWorkbook.Worksheets("EDR")
        For Each cell In Range("A113")
            If cell = strNo Then Exit Sub
                Next cell
                Call EDRMouseMessage
    End With

End Sub

Function EDRMouseMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have ordered a Mouse.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You do not need to order a Mouse when you order a Workstation."
    Msg = Msg & vbNewLine
    Msg = Msg & "The Mouse is part of the Workstation kit."
    Msg = Msg & vbNewLine & vbNewLine
    Msg = Msg & "Would you like to remove the Mouse from your order?"
    Title = "EDR"
    Config = vbYesNo + vbInformation + vbDefaultButton1
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("EDR").Select

    'Change the Mouse selection to No.
    If ExcelBox = vbYes Then Call EDRMouseYes

End Function

Function EDRMouseYes()
'Created for Pason by Dragon Wood (August 2015)

    With ActiveWorkbook.Worksheets("EDR")
        .Range("A113").Value = strNo
        .Range("D113").ClearContents
    End With

End Function

Private Sub EDRMousePadCheck()

    'Check to see if a Workstation is ordered.
    If ActiveWorkbook.Worksheets("EDR").Range("A27").Value = strNo Then Exit Sub

    'Check for a Mouse Pad to be ordered.
    With ActiveWorkbook.Worksheets("EDR")
        For Each cell In Range("A114")
            If cell = strNo Then Exit Sub
                Next cell
                Call EDRMousePadMessage
    End With

End Sub

Function EDRMousePadMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have ordered a Mouse Pad.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You do not need to order a Mouse Pad when you order a Workstation."
    Msg = Msg & vbNewLine
    Msg = Msg & "The Mouse Pad is part of the Workstation kit."
    Msg = Msg & vbNewLine & vbNewLine
    Msg = Msg & "Would you like to remove the Mouse Pad from your order?"
    Title = "EDR"
    Config = vbYesNo + vbInformation + vbDefaultButton1
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("EDR").Select

    'Change the MousePad selection to No.
    If ExcelBox = vbYes Then Call EDRMousePadYes

End Function

Function EDRMousePadYes()
'Created for Pason by Dragon Wood (August 2015)

    With ActiveWorkbook.Worksheets("EDR")
        .Range("A114").Value = strNo
        .Range("D114").ClearContents
    End With

End Function

Private Sub EDRPowerCordCheck()

    'Check to see if a Workstation is ordered.
    If ActiveWorkbook.Worksheets("EDR").Range("A27").Value = strNo Then Exit Sub

    'Check for a Power Cord to be ordered.
    With ActiveWorkbook.Worksheets("EDR")
        For Each cell In Range("A115")
            If cell = strNo Then Exit Sub
                Next cell
                Call EDRPowerCordMessage
    End With

End Sub

Function EDRPowerCordMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have ordered a Power Cord.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You do not need to order a power cord when you order a Workstation."
    Msg = Msg & vbNewLine
    Msg = Msg & "The power cord is part of the Workstation kit."
    Msg = Msg & vbNewLine & vbNewLine
    Msg = Msg & "Would you like to remove the power cord from your order?"
    Title = "EDR"
    Config = vbYesNo + vbInformation + vbDefaultButton1
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("EDR").Select

    'Change the Power Cord selection to No.
    If ExcelBox = vbYes Then Call EDRPowerCordYes

End Function

Function EDRPowerCordYes()
'Created for Pason by Dragon Wood (August 2015)

    With ActiveWorkbook.Worksheets("EDR")
        .Range("A115").Value = strNo
        .Range("D115").ClearContents
    End With

End Function

Private Sub EDRSurgeSupressorCheck()

    'Check to see if a Workstation is ordered.
    If ActiveWorkbook.Worksheets("EDR").Range("A27").Value = strNo Then Exit Sub

    'Check for a Surge Supressor to be ordered.
    With ActiveWorkbook.Worksheets("EDR")
        For Each cell In Range("A89")
            If cell = strNo Then Exit Sub
                Next cell
                Call EDRSurgeSupressorMessage
    End With

End Sub

Function EDRSurgeSupressorMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have ordered a Surge Supressor.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You do not need to order a Surge Supressor when you order a Workstation."
    Msg = Msg & vbNewLine
    Msg = Msg & "The Surge Supressor is part of the Workstation kit."
    Msg = Msg & vbNewLine & vbNewLine
    Msg = Msg & "Would you like to remove the Surge Supressor from your order?"
    Title = "EDR"
    Config = vbYesNo + vbInformation + vbDefaultButton1
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("EDR").Select

    'Change the Surge Supressor selection to No.
    If ExcelBox = vbYes Then Call EDRSurgeSupressorYes

End Function

Function EDRSurgeSupressorYes()
'Created for Pason by Dragon Wood (August 2015)

    With ActiveWorkbook.Worksheets("EDR")
        .Range("A89").Value = strNo
        .Range("D89").ClearContents
    End With

End Function

Private Sub DHCWallMountCheck()
'Created for Pason by Dragon Wood (September 2015).

    If ActiveWorkbook.Worksheets("EDR").Range("A81").Value = strYes Then
        ActiveWorkbook.Worksheets("EDR").Range("D81").Value = 2
    End If

End Sub

Private Sub EDRTargetRingCheck()

    'Check for a Target Ring to be ordered.
    With ActiveWorkbook.Worksheets("EDR")
        For Each cell In Range("A96:A106")
            If cell = strYes Then Exit Sub
                Next cell
                Call EDRTargetRingMessage
    End With

End Sub

Function EDRTargetRingMessage()
'Created for Pason by Dragon Wood (October 2015)
'Warns the user they have not ordered a Target Ring.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Target Ring."
    Title = "EDR"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("EDR").Select

End Function

Private Sub EDRRotaryTargetCheck()

    'Check for a Rotary Target to be ordered.
    With ActiveWorkbook.Worksheets("EDR")
        For Each cell In Range("A107:A108")
            If cell = strYes Then Exit Sub
                Next cell
                Call EDRRotaryTargetMessage
    End With

End Sub

Function EDRRotaryTargetMessage()
'Created for Pason by Dragon Wood (October 2015)
'Warns the user they have not ordered a Rotary Target.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Rotary Target."
    Title = "EDR"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("EDR").Select

End Function

Private Sub EDRDepthBracketABCheck()
'Created for Pason by Dragon Wood (October 2015)
'If a depth sensor is ordered, it checks to see if the correct depth sensor bracket is ordered.

    If ActiveWorkbook.Worksheets("EDR").Range("A28").Value = strYes Then Exit Sub

    'Check for a depth sensor bracket to be ordered.
    With ActiveWorkbook.Worksheets("EDR")
        For Each cell In Range("A41:A42")
            If cell = strYes Then Call EDRDepthBracketABMessage
                Next cell

    End With
End Sub

Function EDRDepthBracketABMessage()
'Created for Pason by Dragon Wood (October 2015)
'Warns the user if they ordered the wrong depth sensor bracket.

    'Declare the variables
    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExclBox As Integer

    'Set the Message Box settings.
    Msg = "You ordered an A/B Depth Sensor, but did not order an A/B Depth Sensor Bracket."
    Msg = Msg & vbNewLine & vbNewLine
    Msg = Msg & "Would you like to add an A/B Depth Sensor Bracket to your order?"
    Title = "EDR Warning!"
    Config = vbYesNo + vbExclamation + vbDefaultButton1
    ExclBox = MsgBox(Msg, Config, Title)

    'Change the depth sensor bracket selection to Yes.
    If ExclBox = vbYes Then Call EDRDepthBracketABYes

    'Verify the user does not want a depth sensor bracket.
    If ExclBox = vbNo Then Call EDRDepthBracketABNo

End Function

Function EDRDepthBracketABYes()
'Created for Pason by Dragon Wood (October 2015)
'Changes an A/B Depth Sensor Bracket selection to Yes.

        With ActiveWorkbook.Worksheets("EDR")
            .Range("A28").Value = strYes
            .Range("D28").Value = 1
            .Range("A29").Value = strNo
            .Range("D29").ClearContents
        End With

End Function

Function EDRDepthBracketABNo()
'Created for Pason by Dragon Wood (October 2015)
'Verifies with the user they do not want a depth sensor bracket.

    'Declare the variables
    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExclBox As Integer

    'Set the Message Box settings.
    Msg = "Are you sure you do not want a depth sensor bracket?"
    Title = "Are You Sure?"
    Config = vbYesNo + vbQuestion + vbDefaultButton2
    ExclBox = MsgBox(Msg, Config, Title)

    'Change the depth sensor bracket selection to Yes.
    If ExclBox = vbNo Then EDRDepthBracketABYes

End Function

Private Sub EDRDepthBracketCCheck()
'Created for Pason by Dragon Wood (October 2015)
'If a depth sensor is ordered, it checks to see if the correct depth sensor bracket is ordered.

    If ActiveWorkbook.Worksheets("EDR").Range("A29").Value = strYes Then Exit Sub

    'Check for a depth sensor bracket to be ordered.
    With ActiveWorkbook.Worksheets("EDR")
        For Each cell In Range("A43:A44")
            If cell = strYes Then Call EDRDepthBracketCMessage
                Next cell

    End With
End Sub

Function EDRDepthBracketCMessage()
'Created for Pason by Dragon Wood (October 2015)
'Warns the user if they ordered the wrong depth sensor bracket.

    'Declare the variables
    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExclBox As Integer

    'Set the Message Box settings.
    Msg = "You ordered a C Depth Sensor, but did not order an C/D/E Depth Sensor Bracket."
    Msg = Msg & vbNewLine & vbNewLine
    Msg = Msg & "Would you like to add an C/D/E Depth Sensor Bracket to your order?"
    Title = "EDR Warning!"
    Config = vbYesNo + vbExclamation + vbDefaultButton1
    ExclBox = MsgBox(Msg, Config, Title)

    'Change the depth sensor bracket selection to Yes.
    If ExclBox = vbYes Then Call EDRDepthBracketCDEYes

    'Verify the user does not want a depth sensor bracket.
    If ExclBox = vbNo Then Call EDRDepthBracketCDENo

End Function

Function EDRDepthBracketCDEYes()
'Created for Pason by Dragon Wood (October 2015)
'Changes an C/D/E Depth Sensor Bracket selection to Yes.

        With ActiveWorkbook.Worksheets("EDR")
            .Range("A29").Value = strYes
            .Range("D29").Value = 1
            .Range("A28").Value = strNo
            .Range("D28").ClearContents
        End With

End Function

Function EDRDepthBracketCDENo()
'Created for Pason by Dragon Wood (October 2015)
'Verifies with the user they do not want a depth sensor bracket.

    'Declare the variables
    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExclBox As Integer

    'Set the Message Box settings.
    Msg = "Are you sure you do not want a depth sensor bracket?"
    Title = "Are You Sure?"
    Config = vbYesNo + vbQuestion + vbDefaultButton2
    ExclBox = MsgBox(Msg, Config, Title)

    'Change the depth sensor bracket selection to Yes.
    If ExclBox = vbNo Then EDRDepthBracketCDEYes

End Function

Private Sub EDRDepthBracketDCheck()
'Created for Pason by Dragon Wood (October 2015)
'If a depth sensor is ordered, it checks to see if the correct depth sensor bracket is ordered.

    If ActiveWorkbook.Worksheets("EDR").Range("A29").Value = strYes Then Exit Sub

    'Check for a depth sensor bracket to be ordered.
    If ActiveWorkbook.Worksheets("EDR").Range("A45").Value = strYes Then EDRDepthBracketDMessage

End Sub

Function EDRDepthBracketDMessage()
'Created for Pason by Dragon Wood (October 2015)
'Warns the user if they ordered the wrong depth sensor bracket.

    'Declare the variables
    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExclBox As Integer

    'Set the Message Box settings.
    Msg = "You ordered a D Depth Sensor, but did not order an C/D/E Depth Sensor Bracket."
    Msg = Msg & vbNewLine & vbNewLine
    Msg = Msg & "Would you like to add an C/D/E Depth Sensor Bracket to your order?"
    Title = "Warning!"
    Config = vbYesNo + vbExclamation + vbDefaultButton1
    ExclBox = MsgBox(Msg, Config, Title)

    'Change the depth sensor bracket selection to Yes.
    If ExclBox = vbYes Then Call EDRDepthBracketCDEYes

    'Verify the user does not want a depth sensor bracket.
    If ExclBox = vbNo Then Call EDRDepthBracketCDENo

End Function

Private Sub EDRDepthBracketECheck()
'Created for Pason by Dragon Wood (October 2015)
'If a depth sensor is ordered, it checks to see if the correct depth sensor bracket is ordered.

    If ActiveWorkbook.Worksheets("EDR").Range("A29").Value = strYes Then Exit Sub

    'Check for a depth sensor bracket to be ordered.
    If ActiveWorkbook.Worksheets("EDR").Range("A46").Value = strYes Then EDRDepthBracketEMessage

End Sub

Function EDRDepthBracketEMessage()
'Created for Pason by Dragon Wood (October 2015)
'Warns the user if they ordered the wrong depth sensor bracket.

    'Declare the variables
    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExclBox As Integer

    'Set the Message Box settings.
    Msg = "You ordered an E Depth Sensor, but did not order an C/D/E Depth Sensor Bracket."
    Msg = Msg & vbNewLine & vbNewLine
    Msg = Msg & "Would you like to add an C/D/E Depth Sensor Bracket to your order?"
    Title = "Warning!"
    Config = vbYesNo + vbExclamation + vbDefaultButton1
    ExclBox = MsgBox(Msg, Config, Title)

    'Change the depth sensor bracket selection to Yes.
    If ExclBox = vbYes Then Call EDRDepthBracketCDEYes

    'Verify the user does not want a depth sensor bracket.
    If ExclBox = vbNo Then Call EDRDepthBracketCDENo

End Function

Private Sub EDRSubUJBPigtailCheck()
'Created for Pason by Dragon Wood (November 2015)
'If a UJB Sub J-Box is ordered, it checks to see if a Sub UJB to UJB Pigtail is ordered.

    If ActiveWorkbook.Worksheets("EDR").Range("A24").Value = strNo Then Exit Sub

    'Check for a Sub UJB to UJB Pigtail to be ordered.
    If ActiveWorkbook.Worksheets("EDR").Range("A76").Value = strYes Then Exit Sub

    'Warn the user if it wasn't.
    Call EDRSubUJBPigtailMessage

End Sub

Function EDRSubUJBPigtailMessage()
'Created for Pason by Dragon Wood (November 2015)
'Warns the user if they ordered a UJB Sub J-Box but did not order a Sub UJB to UJB Pigtail

    'Declare the variables
    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExclBox As Integer

    'Set the Message Box settings.
    Msg = "You ordered a UJB Sub J-Box, but did not order a Sub UJB to UJB Pigtail."
    Msg = Msg & vbNewLine & vbNewLine
    Msg = Msg & "Would you like to add a Sub UJB to UJB Pigtail to your order?"
    Title = "Warning!"
    Config = vbYesNo + vbExclamation + vbDefaultButton1
    ExclBox = MsgBox(Msg, Config, Title)

    'Change the Sub UJB to UJB Pigtail selection to Yes.
    If ExclBox = vbYes Then Call EDRSubUJBPigtailYes

    'Verify the user does not want a Sub UJB to UJB Pigtail.
    If ExclBox = vbNo Then Call EDRSubUJBPigtailNo

End Function

Function EDRSubUJBPigtailYes()
'Created for Pason by Dragon Wood (November 2015)
'Changes a montior selection to Yes

        With ActiveWorkbook.Worksheets("EDR")
            .Range("A76").Value = strYes
            .Range("D24").Copy
            .Range("D76").PasteSpecial xlPasteValues
        End With

End Function

Function EDRSubUJBPigtailNo()
'Created for Pason by Dragon Wood (November 2015)
'Verifies with the user they do not want a Sub UJB to UJB Pigtail.

    'Declare the variables
    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExclBox As Integer

    'Set the Message Box settings.
    Msg = "Are you sure you do not want a Sub UJB to UJB Pigtail?"
    Title = "Are You Sure?"
    Config = vbYesNo + vbQuestion + vbDefaultButton2
    ExclBox = MsgBox(Msg, Config, Title)

    'Change the Sub UJB to UJB Pigtail selection to Yes.
    If ExclBox = vbNo Then EDRSubUJBPigtailYes

End Function

Private Sub EDRSubUJBCableCheck()

    If ActiveWorkbook.Worksheets("EDR").Range("A24").Value = strNo Then Exit Sub

    'Check for a Sub UJB Sensor Cable to be ordered.
    With ActiveWorkbook.Worksheets("EDR")
        For Each cell In Range("A70:A74")
            If cell = strYes Then Exit Sub
                Next cell
                Call EDRSubUJBCableMessage
    End With

End Sub

Function EDRSubUJBCableMessage()
'Created for Pason by Dragon Wood (November 2015)
'Warns the user they have not ordered a Sub UJB Sensor Cable.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Sub UJB Sensor Cable."
    Title = "EDR"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("EDR").Select

End Function

Function EDRCopyData()
    Dim wkShtMaster As Worksheet
    Dim varSheets As Variant
    Dim lngRow As Long
    Dim lngMatch As Long
    Dim lngIncrement As Long
    Dim shpIcon As Shape
     
    Set wkShtMaster = Worksheets("Master DataList")
         
        With Worksheets("EDR")
             
            lngRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            For lngIncrement = 2 To lngRow
                 
                lngMatch = 0
                On Error Resume Next
                lngMatch = Application.Match(.Cells(lngIncrement, "C"), wkShtMaster.Columns("A"), 0)
                On Error GoTo 0
                If lngMatch > 0 Then
                     
                    .Cells(lngIncrement, "B").Value = wkShtMaster.Cells(lngMatch, "B").Value
                    Set shpIcon = GetShape(wkShtMaster, .Cells(lngMatch, "C"))
                    If Not shpIcon Is Nothing Then
                        .Hyperlinks.Add Anchor:=.Cells(lngIncrement, "E"), Address:=shpIcon.Hyperlink.Address, SubAddress:="", TextToDisplay:="Image"
                    End If
                End If
            Next lngIncrement
             
        End With
End Function
