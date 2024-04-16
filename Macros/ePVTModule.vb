Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Public Sub ePVTReset(control As IRibbonControl)
'Created for Pason by Dragon Wood (August 2015)
'Resets the ePVT page so it can be used again.

    Application.ScreenUpdating = False
    
    'Clear all data on the page.
    Call ePVTClear
'    Call ePVTCopyData
    
    'Close the sheet.
    With ActiveWorkbook.Worksheets("System Selection")
        .ePVTBox.Value = strNo
    End With
    
    'Reopen the sheet to refill the required parts.
    With ActiveWorkbook.Worksheets("System Selection")
        .ePVTBox.Value = strYes
    End With
    
    Application.ScreenUpdating = True
    Sheets("System Selection").Select

End Sub

Function ePVTUnhide()
'Created for Pason by Dragon Wood (August 2015)
'Automatically sets the Required equipment to "Yes" and sets the default value for each. These can of course be changed.

    'Ensure the ePVT sheet is selected.
    With ActiveWorkbook.Worksheets("ePVT")
    
        'Set the Required equipment to Yes.
        .Range("A2:A8").Value = strYes
    
        'Set the default quantity for each item on the list
        .Range("D2:D8").Value = 1
    End With
End Function

Function ePVTClear()
'Created for Pason by Dragon Wood (August 2015)
'Resets the ePVT Page to the original unused state.

    Application.ScreenUpdating = False
    Application.EnableEvents = False
   
    'Reset the Yes/No field to No on the ePVT Sheet and clear the Quantity Needed Column
    With ActiveWorkbook.Worksheets("ePVT")
        .Visible = True
        .Range("A2:A53").Value = strNo
        .Range("D2:D53").ClearContents
    End With

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Application.GoTo Sheets("ePVT").Range("A1"), True
    Application.GoTo Sheets("System Selection").Range("A1"), True

End Function

Function ePVTUJB()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they cannot install an ePVT without a UJB and asks if they want to add a UJB to the order.

    'Declare the variables
    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExclBox As Integer
        
    'Set the Message Box settings.
    Msg = "You cannot install an ePVT without a UJB."
    Msg = Msg & vbNewLine & vbNewLine
    Msg = Msg & "Would you like to add the UJB to your order?"
    Title = "Warning!"
    Config = vbYesNo + vbExclamation + vbDefaultButton1
    ExclBox = MsgBox(Msg, Config, Title)
    
    'Activate the UJB box if Yes.
    If ExclBox = vbYes Then UJBYes
    
    'Close the ePVT sheet if No.
    If ExclBox = vbNo Then ePVTUJBNo

End Function

Function ePVTUJBNo()
'Created for Pason by Dragon Wood (August 2015)
'Activates the UJB field if not already active

    ActiveWorkbook.Worksheets("System Selection").ePVTBox.Value = strNo
    ActiveWorkbook.Worksheets("System Selection").Select

End Function

Public Sub ePVTCheck()
'Created for Pason by Dragon Wood (August 2015)
'Checks to make sure that the items that are required, but need a choice are chosen.

    Call ePVTMudProbeBracketCheck
    Call ePVTRadarProbeBracketCheck
    Call ePVTProbeToolCheck
    Call ePVTMudProbeCheck
    Call ePVTRadarProbeCheck
    Call ePVTProbeFloatCheck
    Call ePVTMountCheck
    Call ePVTVegaConnectCheck

End Sub

Private Sub ePVTMudProbeBracketCheck()

    'Check to see if the system is expecting Mud Probes.
    If ActiveWorkbook.Worksheets("System Selection").ProbeBox.Value = strPrbR Then Exit Sub

    'Check for a Mud Probe Bracket to be ordered.
    With ActiveWorkbook.Worksheets("ePVT")
        For Each cell In Range("A9")
            If cell = strYes Then Exit Sub
                Next cell
                Call ePVTMudProbeBracketMessage
    End With

End Sub

Function ePVTMudProbeBracketMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a Mud Probe Bracket.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Mud Probe Bracket."
    Title = "ePVT"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("ePVT").Select

End Function

Private Sub ePVTRadarProbeBracketCheck()

    'Check to see if the system is expecting Radar Probes.
    If ActiveWorkbook.Worksheets("System Selection").ProbeBox.Value = strPrbM Then Exit Sub

    'Check for a Radar Probe Bracket to be ordered.
    With ActiveWorkbook.Worksheets("ePVT")
        For Each cell In Range("A10")
            If cell = strYes Then Exit Sub
                Next cell
                Call ePVTRadarProbeBracketMessage
    End With

End Sub

Function ePVTRadarProbeBracketMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a Radar Probe Bracket.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Radar Probe Bracket."
    Title = "ePVT"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("ePVT").Select

End Function

Private Sub ePVTProbeToolCheck()

    'Check to see if the system is expecting Mud Probes.
    If ActiveWorkbook.Worksheets("System Selection").ProbeBox.Value = strPrbR Then Exit Sub

    'Check for a Mud Probe Tool to be ordered.
    With ActiveWorkbook.Worksheets("ePVT")
        For Each cell In Range("A11")
            If cell = strYes Then Exit Sub
                Next cell
                Call ePVTProbeToolMessage
    End With

End Sub

Function ePVTProbeToolMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a Mud Probe Tool.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Mud Probe Opening Tool."
    Title = "ePVT"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("ePVT").Select

End Function

Private Sub ePVTMudProbeCheck()

    'Check to see if the system is expecting Mud Probes.
    If ActiveWorkbook.Worksheets("System Selection").ProbeBox.Value = strPrbR Then Exit Sub

    'Check for a Mud Probe Bracket to be ordered.
    With ActiveWorkbook.Worksheets("ePVT")
        For Each cell In Range("A12:A16")
            If cell = strYes Then Exit Sub
                Next cell
                Call ePVTMudProbeMessage
    End With

End Sub

Function ePVTMudProbeMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a Mud Probe.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Mud Probe."
    Title = "ePVT"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("ePVT").Select

End Function

Private Sub ePVTRadarProbeCheck()

    'Check to see if the Radar Probe was expected.
    If ActiveWorkbook.Worksheets("System Selection").ProbeBox.Value = strPrbM Then Exit Sub

    'Check for a Mud Probe Bracket to be ordered.
    With ActiveWorkbook.Worksheets("ePVT")
        For Each cell In Range("A17")
            If cell = strYes Then Exit Sub
                Next cell
                Call ePVTRadarProbeMessage
    End With

End Sub

Function ePVTRadarProbeMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a Radar Probe.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Radar Probe."
    Title = "ePVT"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("ePVT").Select

End Function

Private Sub ePVTProbeFloatCheck()

    'Check to see if the system is expecting Mud Probes.
    If ActiveWorkbook.Worksheets("System Selection").ProbeBox.Value = strPrbR Then Exit Sub

    'Check for a Mud Probe Float to be ordered.
    With ActiveWorkbook.Worksheets("ePVT")
        For Each cell In Range("A18")
            If cell = strYes Then Exit Sub
                Next cell
                Call ePVTProbeFloatMessage
    End With

End Sub

Function ePVTProbeFloatMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a Mud Probe Float.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Mud Probe Float."
    Title = "ePVT"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("ePVT").Select

End Function

Private Sub ePVTMountCheck()

    'Check for an ePVT Mount to be ordered.
    With ActiveWorkbook.Worksheets("ePVT")
        For Each cell In Range("A19:A21")
            If cell = strYes Then Exit Sub
                Next cell
                Call ePVTMountMessage
    End With

End Sub

Function ePVTMountMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered an ePVT Mount.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a mount for the ePVT."
    Title = "ePVT"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("ePVT").Select

End Function

Private Sub ePVTVegaConnectCheck()

    'Check to see if the Radar Probe was expected.
    If ActiveWorkbook.Worksheets("System Selection").ProbeBox.Value = strPrbM Then Exit Sub
    
    'Check for a VegaConnect Tool to be ordered.
    With ActiveWorkbook.Worksheets("ePVT")
        For Each cell In Range("A22")
            If cell = strYes Then Exit Sub
                Next cell
                Call ePVTVegaConnectMessage
    End With

End Sub

Function ePVTVegaConnectMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a VegaConnect Tool.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a VegaConnect Tool for the radar probes."
    Title = "ePVT"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("ePVT").Select

End Function

Function ePVTCopyData()
    Dim wkShtMaster As Worksheet
    Dim varSheets As Variant
    Dim lngRow As Long
    Dim lngMatch As Long
    Dim lngIncrement As Long
    Dim shpIcon As Shape
     
    Set wkShtMaster = Worksheets("Master DataList")
         
        With Worksheets("ePVT")
             
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
