Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Public Sub WorkstationsReset(control As IRibbonControl)
'Created for Pason by Dragon Wood (August 2015)
'Resets the Workstations page so it can be used again.

    Application.ScreenUpdating = False
    
    'Clear all data on the page.
    Call WorkstationsClear
'    Call WorkstationsCopyData
    
    'Close the sheet.
    With ActiveWorkbook.Worksheets("System Selection")
        .WorkstationsBox.Value = strNo
    End With
    
    'Reopen the sheet to refill the required parts.
    With ActiveWorkbook.Worksheets("System Selection")
        .WorkstationsBox.Value = strYes
    End With
    
    Application.ScreenUpdating = True
    Sheets("System Selection").Select

End Sub

Function WorkstationsUnhide()
'Created for Pason by Dragon Wood (August 2015)
'Automatically sets the Required equipment to "Yes" and sets the default value for each. These can of course be changed.

    'Ensure the PVT sheet is selected.
    With ActiveWorkbook.Worksheets("Workstations")
    
        'Set the Required equipment to Yes.
        .Range("A2").Value = strYes
    
    End With
    
End Function

Function WorkstationsClear()
'Created for Pason by Dragon Wood (August 2015)
'Resets the Workstations Page to the original unused state.

    Application.ScreenUpdating = False
    Application.EnableEvents = False
   
    'Reset the Yes/No field to No on the Workstations Sheet and clear the Quantity Needed Column
    With ActiveWorkbook.Worksheets("Workstations")
        .Visible = True
        .Range("A2:A25").Value = strNo
        .Range("D2:D25").ClearContents
    End With

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Application.GoTo Sheets("Workstations").Range("A1"), True
    Application.GoTo Sheets("System Selection").Range("A1"), True

End Function

Public Sub WorkstationsCheck()
'Created for Pason by Dragon Wood (August 2015)
'Checks to make sure that the items that are required, but need a choice are chosen.

    Call WSCommCableCheck
    Call WSMonitorCheck
    Call WSKeyboardCheck
    Call WSMouseCheck
    Call WSMousePadCheck
    Call WSPowerCordCheck
    Call WSSurgeSupressorCheck

End Sub

Private Sub WSCommCableCheck()

    'Check for a Comm Cable to be ordered.
    With ActiveWorkbook.Worksheets("Workstations")
        For Each cell In Range("A3:A4")
            If cell = strYes Then Exit Sub
                Next cell
                Call WSCommCableMessage
    End With

End Sub

Function WSCommCableMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a Comm Cable.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a RigComm to Workstation Cable."
    Title = "Workstations"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("Workstations").Select

End Function

Private Sub WSKeyboardCheck()

    'Check to see if a Workstation is ordered.
    If ActiveWorkbook.Worksheets("Workstations").Range("A2").Value = strNo Then Exit Sub

    'Check for a Keyboard to be ordered.
    With ActiveWorkbook.Worksheets("Workstations")
        For Each cell In Range("A22")
            If cell = strNo Then Exit Sub
                Next cell
                Call WSKeyboardMessage
    End With

End Sub

Function WSKeyboardMessage()
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
    Title = "Workstations"
    Config = vbYesNo + vbInformation + vbDefaultButton1
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("Workstations").Select

    'Change the Keyboard selection to No.
    If ExcelBox = vbYes Then Call WSKeyboardYes
    
End Function

Function WSKeyboardYes()

    With ActiveWorkbook.Worksheets("Workstations")
        .Range("A22").Value = strNo
        .Range("D22").ClearContents
    End With
    
End Function

Private Sub WSMonitorCheck()
'Created for Pason by Dragon Wood (August 2015)
'If a workstation is ordered, it checks to see if a monitor is ordered.
    
    'Check to see if a workstation is ordered.
    If ActiveWorkbook.Worksheets("Workstations").Range("A2").Value = strNo Then Exit Sub

    'Check for a monitor to be ordered.
    With ActiveWorkbook.Worksheets("Workstations")
        For Each cell In Range("A5:A8")
            If cell = strYes Then Exit Sub
                Next cell
                Call WSMonitorMessage
    End With
End Sub

Function WSMonitorMessage()
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
    If ExclBox = vbYes Then Call WSMonitorYes
    
    'Verify the user does not want a monitor.
    If ExclBox = vbNo Then Call WSMonitorNo

End Function

Function WSMonitorYes()
'Created for Pason by Dragon Wood (August 2015)
'Changes a montior selection to Yes

        With ActiveWorkbook.Worksheets("Workstations")
            .Range("A8").Value = strYes
            .Range("D2").Copy
            .Range("D8").PasteSpecial xlPasteValues
        End With

End Function

Function WSMonitorNo()
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
    If ExclBox = vbNo Then WSMonitorYes

End Function

Private Sub WSMouseCheck()

    'Check to see if a Workstation is ordered.
    If ActiveWorkbook.Worksheets("Workstations").Range("A2").Value = strNo Then Exit Sub

    'Check for a Mouse to be ordered.
    With ActiveWorkbook.Worksheets("Workstations")
        For Each cell In Range("A23")
            If cell = strNo Then Exit Sub
                Next cell
                Call WSMouseMessage
    End With

End Sub

Function WSMouseMessage()
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
    Title = "Workstations"
    Config = vbYesNo + vbInformation + vbDefaultButton1
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("Workstations").Select

    'Change the Mouse selection to No.
    If ExcelBox = vbYes Then Call WSMouseYes
    
End Function

Function WSMouseYes()

    With ActiveWorkbook.Worksheets("Workstations")
        .Range("A23").Value = strNo
        .Range("D23").ClearContents
    End With
    
End Function

Private Sub WSMousePadCheck()

    'Check to see if a Workstation is ordered.
    If ActiveWorkbook.Worksheets("Workstations").Range("A2").Value = strNo Then Exit Sub

    'Check for a MousePad to be ordered.
    With ActiveWorkbook.Worksheets("Workstations")
        For Each cell In Range("A24")
            If cell = strNo Then Exit Sub
                Next cell
                Call WSMousePadMessage
    End With

End Sub

Function WSMousePadMessage()
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
    Title = "Workstations"
    Config = vbYesNo + vbInformation + vbDefaultButton1
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("Workstations").Select

    'Change the MousePad selection to No.
    If ExcelBox = vbYes Then Call WSMousePadYes
    
End Function

Function WSMousePadYes()

    With ActiveWorkbook.Worksheets("Workstations")
        .Range("A24").Value = strNo
        .Range("D24").ClearContents
    End With
    
End Function

Private Sub WSPowerCordCheck()

    'Check to see if a Workstation is ordered.
    If ActiveWorkbook.Worksheets("Workstations").Range("A2").Value = strNo Then Exit Sub

    'Check for a PowerCord to be ordered.
    With ActiveWorkbook.Worksheets("Workstations")
        For Each cell In Range("A20")
            If cell = strNo Then Exit Sub
                Next cell
                Call WSPowerCordMessage
    End With

End Sub

Function WSPowerCordMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have ordered a Power Cord.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You do not need to order a Power Cord when you order a Workstation."
    Msg = Msg & vbNewLine
    Msg = Msg & "The Power Cord is part of the Workstation kit."
    Msg = Msg & vbNewLine & vbNewLine
    Msg = Msg & "Would you like to remove the Power Cord from your order?"
    Title = "Workstations"
    Config = vbYesNo + vbInformation + vbDefaultButton1
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("Workstations").Select

    'Change the PowerCord selection to No.
    If ExcelBox = vbYes Then Call WSPowerCordYes
    
End Function

Function WSPowerCordYes()

    With ActiveWorkbook.Worksheets("Workstations")
        .Range("A20").Value = strNo
        .Range("D20").ClearContents
    End With
    
End Function

Private Sub WSSurgeSupressorCheck()

    'Check to see if a Workstation is ordered.
    If ActiveWorkbook.Worksheets("Workstations").Range("A2").Value = strNo Then Exit Sub

    'Check for a Surge Supressor to be ordered.
    With ActiveWorkbook.Worksheets("Workstations")
        For Each cell In Range("A25")
            If cell = strNo Then Exit Sub
                Next cell
                Call WSSurgeSupressorMessage
    End With

End Sub

Function WSSurgeSupressorMessage()
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
    Title = "Workstations"
    Config = vbYesNo + vbInformation + vbDefaultButton1
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("Workstations").Select

    'Change the Surge Supressor selection to No.
    If ExcelBox = vbYes Then Call WSSurgeSupressorYes
    
End Function

Function WSSurgeSupressorYes()

    With ActiveWorkbook.Worksheets("Workstations")
        .Range("A25").Value = strNo
        .Range("D25").ClearContents
    End With
    
End Function

Function WorkstationsCopyData()
    Dim wkShtMaster As Worksheet
    Dim varSheets As Variant
    Dim lngRow As Long
    Dim lngMatch As Long
    Dim lngIncrement As Long
    Dim shpIcon As Shape
     
    Set wkShtMaster = Worksheets("Master DataList")
         
        With Worksheets("Workstations")
             
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
