Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Public Sub PVTReset(control As IRibbonControl)
'Created for Pason by Dragon Wood (August 2015)
'Resets the PVT page so it can be used again.

    Application.ScreenUpdating = False
    
    'Clear all data on the page.
    Call PVTClear
'    Call PVTCopyData
    
    'Close the sheet.
    With ActiveWorkbook.Worksheets("System Selection")
        .PVTBox.Value = strNo
    End With
    
    'Reopen the sheet to refill the required parts.
    With ActiveWorkbook.Worksheets("System Selection")
        .PVTBox.Value = strYes
    End With
    
    Application.ScreenUpdating = True
    Sheets("System Selection").Select

End Sub

Function PVTUnhide()
'Created for Pason by Dragon Wood (August 2015)
'Automatically sets the Required equipment to "Yes" and sets the default value for each. These can of course be changed.

    'Ensure the PVT sheet is selected.
    With ActiveWorkbook.Worksheets("PVT")
    
        'Set the Required equipment to Yes.
        .Range("A2:A7").Value = strYes
    
        'Set the default quantity for each item on the list
        .Range("D2:D7").Value = 1
    End With
    
End Function

Function PVTClear()
'Created for Pason by Dragon Wood (August 2015)
'Resets the PVT Page to the original unused state.

    Application.ScreenUpdating = False
    Application.EnableEvents = False
   
    'Reset the Yes/No field to No on the PVT Sheet and clear the Quantity Needed Column
    With ActiveWorkbook.Worksheets("PVT")
        .Visible = True
        .Range("A2:A58").Value = strNo
        .Range("D2:D58").ClearContents
    End With

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Application.GoTo Sheets("PVT").Range("A1"), True
    Application.GoTo Sheets("System Selection").Range("A1"), True

End Function

Public Sub PVTCheck()
'Created for Pason by Dragon Wood (August 2015)
'Checks to make sure that the items that are required, but need a choice are chosen.

    Call PVTMudProbeBracketCheck
    Call PVTRadarProbeBracketCheck
    Call PVTMainCableCheck
    Call PVTProbeToolCheck
    Call PVTJBoxCheck
    Call PVTMudProbeCheck
    Call PVTRadarProbeCheck
    Call PVTProbeFloatCheck
    Call PVTUJBCheck
    Call PVTVegaConnectCheck
    Call PVTMountCheck

End Sub

Private Sub PVTMudProbeBracketCheck()

    'Check to see if the system is expecting Mud Probes.
    If ActiveWorkbook.Worksheets("System Selection").ProbeBox.Value = strPrbR Then Exit Sub

    'Check for a Mud Probe Bracket to be ordered.
    With ActiveWorkbook.Worksheets("PVT")
        For Each cell In Range("A8")
            If cell = strYes Then Exit Sub
                Next cell
                Call PVTMudProbeBracketMessage
    End With

End Sub

Function PVTMudProbeBracketMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a Mud Probe Bracket.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Mud Probe Bracket."
    Title = "PVT"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("PVT").Select

End Function

Private Sub PVTRadarProbeBracketCheck()

    'Check to see if the system is expecting Radar Probes.
    If ActiveWorkbook.Worksheets("System Selection").ProbeBox.Value = strPrbM Then Exit Sub

    'Check for a Radar Probe Bracket to be ordered.
    With ActiveWorkbook.Worksheets("PVT")
        For Each cell In Range("A9")
            If cell = strYes Then Exit Sub
                Next cell
                Call PVTRadarProbeBracketMessage
    End With

End Sub

Function PVTRadarProbeBracketMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a Radar Probe Bracket.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Radar Probe Bracket."
    Title = "PVT"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("PVT").Select

End Function

Private Sub PVTMainCableCheck()

    'Check to see if the system is expecting a UJB.
    If ActiveWorkbook.Worksheets("System Selection").UJBBox.Value = strYes Then Exit Sub

    'Check for a PVT Main Cable to be ordered.
    With ActiveWorkbook.Worksheets("PVT")
        For Each cell In Range("A10:A16")
            If cell = strYes Then Exit Sub
                Next cell
                Call PVTMainCableMessage
    End With

End Sub

Function PVTMainCableMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a PVT Main Cable.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a PVT Main Cable."
    Title = "PVT"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("PVT").Select

End Function

Private Sub PVTProbeToolCheck()

    'Check to see if the system is expecting Mud Probes.
    If ActiveWorkbook.Worksheets("System Selection").ProbeBox.Value = strPrbR Then Exit Sub

    'Check for a Mud Probe Tool to be ordered.
    With ActiveWorkbook.Worksheets("PVT")
        For Each cell In Range("A17")
            If cell = strYes Then Exit Sub
                Next cell
                Call PVTProbeToolMessage
    End With

End Sub

Function PVTProbeToolMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a Mud Probe Tool.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Mud Probe Opening Tool."
    Title = "PVT"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("PVT").Select

End Function

Private Sub PVTJBoxCheck()

    'Check to see if the system is expecting a UJB.
    If ActiveWorkbook.Worksheets("System Selection").UJBBox.Value = strYes Then Exit Sub

    'Check for a PVT J-Box to be ordered.
    With ActiveWorkbook.Worksheets("PVT")
        For Each cell In Range("A18:A19")
            If cell = strYes Then Exit Sub
                Next cell
                Call PVTJBoxMessage
    End With

End Sub

Function PVTJBoxMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a PVT J-Box.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a PVT J-Box."
    Title = "PVT"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("PVT").Select

End Function

Private Sub PVTMudProbeCheck()

    'Check to see if the system is expecting Mud Probes.
    If ActiveWorkbook.Worksheets("System Selection").ProbeBox.Value = strPrbR Then Exit Sub

    'Check for a Mud Probe to be ordered.
    With ActiveWorkbook.Worksheets("PVT")
        For Each cell In Range("A20:A24")
            If cell = strYes Then Exit Sub
                Next cell
                Call PVTMudProbeMessage
    End With

End Sub

Function PVTMudProbeMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a Mud Probe.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Mud Probe."
    Title = "PVT"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("PVT").Select

End Function

Private Sub PVTRadarProbeCheck()

    'Check to see if the Radar Probe was expected.
    If ActiveWorkbook.Worksheets("System Selection").ProbeBox.Value = strPrbM Then Exit Sub

    'Check for a Mud Probe Bracket to be ordered.
    With ActiveWorkbook.Worksheets("PVT")
        For Each cell In Range("A25")
            If cell = strYes Then Exit Sub
                Next cell
                Call PVTRadarProbeMessage
    End With

End Sub

Function PVTRadarProbeMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a Radar Probe.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Radar Probe."
    Title = "PVT"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("PVT").Select

End Function

Private Sub PVTProbeFloatCheck()

    'Check to see if the system is expecting Mud Probes.
    If ActiveWorkbook.Worksheets("System Selection").ProbeBox.Value = strPrbR Then Exit Sub

    'Check for a Mud Probe Float to be ordered.
    With ActiveWorkbook.Worksheets("PVT")
        For Each cell In Range("A26")
            If cell = strYes Then Exit Sub
                Next cell
                Call PVTProbeFloatMessage
    End With

End Sub

Function PVTProbeFloatMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a Mud Probe Float.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Mud Probe Float."
    Title = "PVT"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("PVT").Select

End Function

Private Sub PVTUJBCheck()

    'Check to see if the system is expecting a UJB.
    If ActiveWorkbook.Worksheets("System Selection").UJBBox.Value = strNo Then Exit Sub

    'Check for a UJB to be ordered.
    With ActiveWorkbook.Worksheets("PVT")
        For Each cell In Range("A27")
            If cell = strYes Then Exit Sub
                Next cell
                Call PVTUJBMessage
    End With

End Sub

Function PVTUJBMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a UJB.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a UJB."
    Title = "PVT"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("PVT").Select

End Function

Private Sub PVTVegaConnectCheck()

    'Check to see if the Radar Probe was expected.
    If ActiveWorkbook.Worksheets("System Selection").ProbeBox.Value = strPrbM Then Exit Sub
    
    'Check for a VegaConnect Tool to be ordered.
    With ActiveWorkbook.Worksheets("PVT")
        For Each cell In Range("A28")
            If cell = strYes Then Exit Sub
                Next cell
                Call PVTVegaConnectMessage
    End With

End Sub

Function PVTVegaConnectMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a VegaConnect Tool.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a VegaConnect Tool for the radar probes."
    Title = "PVT"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("PVT").Select

End Function

Private Sub PVTMountCheck()

    'Check for an PVT Mount to be ordered.
    With ActiveWorkbook.Worksheets("PVT")
        For Each cell In Range("A53:A56")
            If cell = strYes Then Exit Sub
                Next cell
                Call PVTMountMessage
    End With

End Sub

Function PVTMountMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a PVT Mount.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a mount for the PVT."
    Title = "PVT"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("PVT").Select

End Function

Function PVTCopyData()
    Dim wkShtMaster As Worksheet
    Dim varSheets As Variant
    Dim lngRow As Long
    Dim lngMatch As Long
    Dim lngIncrement As Long
    Dim shpIcon As Shape
     
    Set wkShtMaster = Worksheets("Master DataList")
         
        With Worksheets("PVT")
             
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
