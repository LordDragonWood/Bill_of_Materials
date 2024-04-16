Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Public Sub ESRReset(control As IRibbonControl)
'Created for Pason by Dragon Wood (October 2015)
'Resets the ESR page so it can be used again.

    Application.ScreenUpdating = False

    'Clear all data on the page.
    Call ESRClear
'    Call ESRCopyData

    'Close the sheet.
    With ActiveWorkbook.Worksheets("System Selection")
        .ESRBox.Value = strNo
    End With

    'Reopen the sheet to refill the required parts.
    With ActiveWorkbook.Worksheets("System Selection")
        .ESRBox.Value = strYes
    End With

    Application.ScreenUpdating = True
    Sheets("System Selection").Select
'    Call BrokenButton
    
End Sub

Function ESRUnhide()
'Created for Pason by Dragon Wood (October 2015)
'Automatically sets the Required equipment to "Yes" and sets the default value for each. These can of course be changed.

    'Ensure the ESR sheet is selected.
    With ActiveWorkbook.Worksheets("ESR")

        'Set the Required equipment to Yes.
        .Range("A2:A27").Value = strYes

        'Set the default quantity for each item on the list
        .Range("D2:D3").Value = 1
        .Range("D4").Value = 3
        .Range("D5:D18").Value = 1
        .Range("D20:D27").Value = 1
    End With

End Function

Function ESRClear()
'Created for Pason by Dragon Wood (October 2015)
'Resets the ESR Page to the original unused state.

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    'Reset the Yes/No field to No on the ESR Sheet and clear the Quantity Needed Column
    With ActiveWorkbook.Worksheets("ESR")
        .Visible = True
        .Range("A2:A139").Value = strNo
        .Range("D2:D139").ClearContents
    End With

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    Application.GoTo Sheets("ESR").Range("A1"), True
    Application.GoTo Sheets("System Selection").Range("A1"), True

End Function

Function ESRUJB()
'Created for Pason by Dragon Wood (October 2015)
'Warns the user they cannot install an ESR without a UJB and asks if they want to add a UJB to the order.

    'Declare the variables
    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExclBox As Integer
        
    'Set the Message Box settings.
    Msg = "You cannot install an ESR without a UJB."
    Msg = Msg & vbNewLine & vbNewLine
    Msg = Msg & "Would you like to add the UJB to your order?"
    Title = "Warning!"
    Config = vbYesNo + vbExclamation + vbDefaultButton1
    ExclBox = MsgBox(Msg, Config, Title)
    
    'Activate the UJB box if Yes.
    If ExclBox = vbYes Then UJBYes
    
    'Close the ePVT sheet if No.
    If ExclBox = vbNo Then ESRUJBNo

End Function

Function ESRUJBNo()
'Created for Pason by Dragon Wood (October 2015)
'Activates the UJB field if not already active

    ActiveWorkbook.Worksheets("System Selection").ESRBox.Value = strNo
    ActiveWorkbook.Worksheets("System Selection").Select

End Function

Public Sub ESRCheck()
'Created for Pason by Dragon Wood (October 2015)
'Checks to make sure that the items that are required, but need a choice are chosen.

    Call ESRDepthBracketABCheck
    Call ESRDepthBracketCCheck
    Call ESRDepthBracketDCheck
    Call ESRDepthBracketECheck
    Call ESRDepthBracketCheck
    Call ESRDepthCableCheck
    Call ESRDepthSensorCheck
    Call ESRSandlineCableCheck
    Call ESREthernetCableCheck
    Call ESRHoseClampCheck
    Call ESRHookloadSensorCheck
    Call ESRRotarySensorCheck
    Call ESRTorqueSensorCheck
    Call ESRTargetRingCheck
    Call ESRRotaryTargetCheck

End Sub

Private Sub ESRDepthBracketCheck()

    'Check for a Depth Sensor Bracket to be ordered.
    With ActiveWorkbook.Worksheets("ESR")
        For Each cell In Range("A28:A29")
            If cell = strYes Then Exit Sub
                Next cell
                Call ESRDepthBracketMessage
    End With

End Sub

Function ESRDepthBracketMessage()
'Created for Pason by Dragon Wood (October 2015)
'Warns the user they have not ordered a Depth Sensor Bracket.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Depth Sensor Bracket."
    Title = "ESR"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("ESR").Select

End Function

Private Sub ESRDepthCableCheck()

    'Check for a Depth Cable to be ordered.
    With ActiveWorkbook.Worksheets("ESR")
        For Each cell In Range("A30:A37")
            If cell = strYes Then Exit Sub
                Next cell
                Call ESRDepthCableMessage
    End With

End Sub

Function ESRDepthCableMessage()
'Created for Pason by Dragon Wood (October 2015)
'Warns the user they have not ordered a Depth Cable.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Depth Cable."
    Title = "ESR"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("ESR").Select

End Function

Private Sub ESRDepthSensorCheck()

    'Check for a Depth Sensor to be ordered.
    With ActiveWorkbook.Worksheets("ESR")
        For Each cell In Range("A57:A62")
            If cell = strYes Then Exit Sub
                Next cell
                Call ESRDepthSensorMessage
    End With

End Sub

Function ESRDepthSensorMessage()
'Created for Pason by Dragon Wood (October 2015)
'Warns the user they have not ordered a Depth Sensor.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Depth Sensor."
    Title = "ESR"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("ESR").Select

End Function

Private Sub ESRSandlineCableCheck()

    'Check for a Sandline Sensor Cable to be ordered.
    With ActiveWorkbook.Worksheets("ESR")
        For Each cell In Range("A38:A46")
            If cell = strYes Then Exit Sub
                Next cell
                Call ESRSandlineCableMessage
    End With

End Sub

Function ESRSandlineCableMessage()
'Created for Pason by Dragon Wood (October 2015)
'Warns the user they have not ordered a Sandline Sensor Cable.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Sandline Sensor Cable."
    Title = "ESR"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("ESR").Select

End Function

Private Sub ESREthernetCableCheck()

    'Check for a Ethernet Cable to be ordered.
    With ActiveWorkbook.Worksheets("ESR")
        For Each cell In Range("A47:A52")
            If cell = strYes Then Exit Sub
                Next cell
                Call ESREthernetCableMessage
    End With

End Sub

Function ESREthernetCableMessage()
'Created for Pason by Dragon Wood (October 2015)
'Warns the user they have not ordered a Ethernet Cable.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order an Ethernet Cable."
    Title = "ESR"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("ESR").Select

End Function

Private Sub ESRHoseClampCheck()

    'Check for a Hose Clamp to be ordered.
    With ActiveWorkbook.Worksheets("ESR")
        For Each cell In Range("A53:A56")
            If cell = strYes Then Exit Sub
                Next cell
                Call ESRHoseClampMessage
    End With

End Sub

Function ESRHoseClampMessage()
'Created for Pason by Dragon Wood (October 2015)
'Warns the user they have not ordered a Hose Clamp.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Hose Clamp."
    Title = "ESR"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("ESR").Select

End Function

Private Sub ESRHookloadSensorCheck()

    'Check for a Hookload Sensor to be ordered.
    With ActiveWorkbook.Worksheets("ESR")
        For Each cell In Range("A63:A65")
            If cell = strYes Then Exit Sub
                Next cell
                Call ESRHookloadSensorMessage
    End With

End Sub

Function ESRHookloadSensorMessage()
'Created for Pason by Dragon Wood (October 2015)
'Warns the user they have not ordered a Hookload Sensor.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Hookload Sensor."
    Title = "ESR"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("ESR").Select

End Function

Private Sub ESRRotarySensorCheck()

    'Check for a Rotary Sensor to be ordered.
    With ActiveWorkbook.Worksheets("ESR")
        For Each cell In Range("A66:A67")
            If cell = strYes Then Exit Sub
                Next cell
                Call ESRRotarySensorMessage
    End With

End Sub

Function ESRRotarySensorMessage()
'Created for Pason by Dragon Wood (October 2015)
'Warns the user they have not ordered a Rotary Sensor.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Rotary Sensor."
    Title = "ESR"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("ESR").Select

End Function

Private Sub ESRTorqueSensorCheck()

    If ActiveWorkbook.Worksheets("ESR").Range("A22").Value = strYes Then Exit Sub

    'Check for a Torque Sensor to be ordered.
    With ActiveWorkbook.Worksheets("ESR")
        For Each cell In Range("A68:A69")
            If cell = strYes Then Exit Sub
                Next cell
                Call ESRTorqueSensorMessage
    End With

End Sub

Function ESRTorqueSensorMessage()
'Created for Pason by Dragon Wood (October 2015)
'Warns the user they have not ordered a Torque Sensor.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Torque Sensor."
    Title = "ESR"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("ESR").Select

End Function

Private Sub ESRTargetRingCheck()

    'Check for a Target Ring to be ordered.
    With ActiveWorkbook.Worksheets("ESR")
        For Each cell In Range("A70:A83")
            If cell = strYes Then Exit Sub
                Next cell
                Call ESRTargetRingMessage
    End With

End Sub

Function ESRTargetRingMessage()
'Created for Pason by Dragon Wood (October 2015)
'Warns the user they have not ordered a Target Ring.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Target Ring."
    Title = "ESR"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("ESR").Select

End Function

Private Sub ESRRotaryTargetCheck()

    'Check for a Rotary Target to be ordered.
    With ActiveWorkbook.Worksheets("ESR")
        For Each cell In Range("A84:A85")
            If cell = strYes Then Exit Sub
                Next cell
                Call ESRRotaryTargetMessage
    End With

End Sub

Function ESRRotaryTargetMessage()
'Created for Pason by Dragon Wood (October 2015)
'Warns the user they have not ordered a Rotary Target.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Rotary Target."
    Title = "ESR"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("ESR").Select

End Function

Private Sub ESRDepthBracketABCheck()
'Created for Pason by Dragon Wood (October 2015)
'If a depth sensor is ordered, it checks to see if the correct depth sensor bracket is ordered.

    If ActiveWorkbook.Worksheets("ESR").Range("A28").Value = strYes Then Exit Sub
    
    'Check for a depth sensor bracket to be ordered.
    With ActiveWorkbook.Worksheets("ESR")
        For Each cell In Range("A57:A58")
            If cell = strYes Then Call ESRDepthBracketABMessage 'Exit Sub
                Next cell
    
    End With
End Sub

Function ESRDepthBracketABMessage()
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
    Title = "Warning!"
    Config = vbYesNo + vbExclamation + vbDefaultButton1
    ExclBox = MsgBox(Msg, Config, Title)
    
    'Change the depth sensor bracket selection to Yes.
    If ExclBox = vbYes Then Call ESRDepthBracketABYes
    
    'Verify the user does not want a depth sensor bracket.
    If ExclBox = vbNo Then Call ESRDepthBracketABNo

End Function

Function ESRDepthBracketABYes()
'Created for Pason by Dragon Wood (October 2015)
'Changes an A/B Depth Sensor Bracket selection to Yes.

        With ActiveWorkbook.Worksheets("ESR")
            .Range("A28").Value = strYes
            .Range("D28").Value = 1
            .Range("A29").Value = strNo
            .Range("D29").ClearContents
        End With

End Function

Function ESRDepthBracketABNo()
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
    If ExclBox = vbNo Then ESRDepthBracketABYes

End Function

Private Sub ESRDepthBracketCCheck()
'Created for Pason by Dragon Wood (October 2015)
'If a depth sensor is ordered, it checks to see if the correct depth sensor bracket is ordered.

    If ActiveWorkbook.Worksheets("ESR").Range("A29").Value = strYes Then Exit Sub
    
    'Check for a depth sensor bracket to be ordered.
    With ActiveWorkbook.Worksheets("ESR")
        For Each cell In Range("A59:A60")
            If cell = strYes Then Call ESRDepthBracketCMessage
                Next cell
    
    End With
End Sub

Function ESRDepthBracketCMessage()
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
    Title = "Warning!"
    Config = vbYesNo + vbExclamation + vbDefaultButton1
    ExclBox = MsgBox(Msg, Config, Title)
    
    'Change the depth sensor bracket selection to Yes.
    If ExclBox = vbYes Then Call ESRDepthBracketCYes
    
    'Verify the user does not want a depth sensor bracket.
    If ExclBox = vbNo Then Call ESRDepthBracketCNo

End Function

Function ESRDepthBracketCYes()
'Created for Pason by Dragon Wood (October 2015)
'Changes an C/D/E Depth Sensor Bracket selection to Yes.

        With ActiveWorkbook.Worksheets("ESR")
            .Range("A29").Value = strYes
            .Range("D29").Value = 1
            .Range("A28").Value = strNo
            .Range("D28").ClearContents
        End With

End Function

Function ESRDepthBracketCNo()
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
    If ExclBox = vbNo Then ESRDepthBracketCYes

End Function

Private Sub ESRDepthBracketDCheck()
'Created for Pason by Dragon Wood (October 2015)
'If a depth sensor is ordered, it checks to see if the correct depth sensor bracket is ordered.

    If ActiveWorkbook.Worksheets("ESR").Range("A29").Value = strYes Then Exit Sub

    'Check for a depth sensor bracket to be ordered.
    If ActiveWorkbook.Worksheets("ESR").Range("A61").Value = strYes Then ESRDepthBracketDMessage
    
End Sub

Function ESRDepthBracketDMessage()
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
    If ExclBox = vbYes Then Call ESRDepthBracketDYes

    'Verify the user does not want a depth sensor bracket.
    If ExclBox = vbNo Then Call ESRDepthBracketDNo

End Function

Function ESRDepthBracketDYes()
'Created for Pason by Dragon Wood (October 2015)
'Changes an C/D/E Depth Sensor Bracket selection to Yes.

        With ActiveWorkbook.Worksheets("ESR")
            .Range("A29").Value = strYes
            .Range("D29").Value = 1
            .Range("A28").Value = strNo
            .Range("D28").ClearContents
        End With

End Function

Function ESRDepthBracketDNo()
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
    If ExclBox = vbNo Then ESRDepthBracketDYes

End Function

Private Sub ESRDepthBracketECheck()
'Created for Pason by Dragon Wood (October 2015)
'If a depth sensor is ordered, it checks to see if the correct depth sensor bracket is ordered.

    If ActiveWorkbook.Worksheets("ESR").Range("A29").Value = strYes Then Exit Sub

    'Check for a depth sensor bracket to be ordered.
    If ActiveWorkbook.Worksheets("ESR").Range("A62").Value = strYes Then ESRDepthBracketEMessage
    
End Sub

Function ESRDepthBracketEMessage()
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
    If ExclBox = vbYes Then Call ESRDepthBracketEYes

    'Verify the user does not want a depth sensor bracket.
    If ExclBox = vbNo Then Call ESRDepthBracketENo

End Function

Function ESRDepthBracketEYes()
'Created for Pason by Dragon Wood (October 2015)
'Changes an C/D/E Depth Sensor Bracket selection to Yes.

        With ActiveWorkbook.Worksheets("ESR")
            .Range("A29").Value = strYes
            .Range("D29").Value = 1
            .Range("A28").Value = strNo
            .Range("D28").ClearContents
        End With

End Function

Function ESRDepthBracketENo()
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
    If ExclBox = vbNo Then ESRDepthBracketEYes

End Function

Function ESRCopyData()
    Dim wkShtMaster As Worksheet
    Dim varSheets As Variant
    Dim lngRow As Long
    Dim lngMatch As Long
    Dim lngIncrement As Long
    Dim shpIcon As Shape
     
    Set wkShtMaster = Worksheets("Master DataList")
         
        With Worksheets("ESR")
             
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
