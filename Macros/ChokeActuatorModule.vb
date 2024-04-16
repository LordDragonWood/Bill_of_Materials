Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Public Sub ChokeActuatorReset(control As IRibbonControl)
'Created for Pason by Dragon Wood (August 2015)
'Resets the Choke Actuator page so it can be used again.

    Application.ScreenUpdating = False
    
    'Clear all data on the page.
    Call ChokeActuatorClear
'    Call ChokeCopyData
    
    'Close the sheet.
    With ActiveWorkbook.Worksheets("System Selection")
        .ChokeBox.Value = strNo
    End With
    
    'Reopen the sheet to refill the required parts.
    With ActiveWorkbook.Worksheets("System Selection")
        .ChokeBox.Value = strYes
    End With
    
    Application.ScreenUpdating = True
    Sheets("System Selection").Select

End Sub

Function ChokeActuatorUnhide()
'Created for Pason by Dragon Wood (August 2015)
'Automatically sets the Required equipment to "Yes" and sets the default value for each. These can of course be changed.

    'Ensure the Choke Actuator sheet is selected.
    With ActiveWorkbook.Worksheets("Choke Actuator")
    
    'Set the Required equipment to Yes.
        .Range("A2:A4").Value = strYes
    
    'Set the default quantity for each item on the list
        .Range("D2").Value = 1
        .Range("D3").Value = 2
        .Range("D4").Value = 1
    End With
    
End Function

Function ChokeActuatorClear()
'Created for Pason by Dragon Wood (August 2015)
'Resets the Choke Actuator Page to the original unused state.

    Application.ScreenUpdating = False
    Application.EnableEvents = False
   
    'Reset the Yes/No field to No on the Choke Actuator Sheet and clear the Quantity Needed Column
    With ActiveWorkbook.Worksheets("Choke Actuator")
        .Visible = True
        .Range("A2:A44").Value = strNo
        .Range("D2:D44").ClearContents
    End With

    Application.ScreenUpdating = True
    Application.EnableEvents = False
    
    Application.GoTo Sheets("Choke Actuator").Range("A1"), True
    Application.GoTo Sheets("System Selection").Range("A1"), True

End Function

Public Sub ChokeActuatorCheck()
'Created for Pason by Dragon Wood (August 2015)
'Checks to make sure that the items that are required, but need a choice are chosen.
    
    Call ChokeActuatorCableCheck
    Call ChokeActuatorCouplerCheck
    Call ChokeCasingCheck
    Call CAUJBSubJCheck
    Call CASubUJBCableCheck

End Sub

Private Sub ChokeActuatorCableCheck()

    'Check for a Choke Actuator Cable to be ordered.
    With ActiveWorkbook.Worksheets("Choke Actuator")
        For Each cell In Range("A5:A7")
            If cell = strYes Then Exit Sub
                Next cell
                Call ChokeActuatorCableMessage
    End With

End Sub

Function ChokeActuatorCableMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a Choke Actuator Cable.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer
    
    Msg = "You did not order a Choke Actuator Cable."
    Title = "Choke Actuator"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("Choke Actuator").Select
    
End Function

Private Sub ChokeActuatorCouplerCheck()

    'Check for a Choke Actuator Coupler to be ordered.
    With ActiveWorkbook.Worksheets("Choke Actuator")
        For Each cell In Range("A8:A9")
            If cell = strYes Then Exit Sub
                Next cell
                Call ChokeActuatorCouplerMessage
    End With

End Sub

Function ChokeActuatorCouplerMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a Choke Actuator Coupler.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer
    
    Msg = "You did not order a Choke Actuator Coupler."
    Title = "Choke Actuator"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("Choke Actuator").Select
    
End Function

Function ChokeCasingCheck()
'Created for Pason by Dragon Wood (September 2015).

    'Warn the user they should not use a Choke without a Casing Pressure Sensor.
    If ActiveWorkbook.Worksheets("System Selection").CasingBox.Value = strNo Then ChokeCasingMessage

End Function

Function ChokeCasingMessage()
'Created for Pason by Dragon Wood (September 2015)
'Warns the user they did not order a Casing Pressure Sensor and asks if they want to add a Casing Pressure Sensor to the order.

    'Declare the variables
    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExclBox As Integer
        
    'Set the Message Box settings.
    Msg = "You did not order a Casing Pressure Sensor."
    Msg = Msg & vbNewLine
    Msg = Msg & "It is not recommended that you install the choke without one."
    Msg = Msg & vbNewLine & vbNewLine
    Msg = Msg & "Would you like to add the Casing Pressure Sensor to your order?"
    Title = "Warning!"
    Config = vbYesNo + vbExclamation + vbDefaultButton1
    ExclBox = MsgBox(Msg, Config, Title)
    
    'Activate the UJB box if Yes.
    If ExclBox = vbYes Then ChokeCasingYes
    
    'Close the ePVT sheet if No.
    If ExclBox = vbNo Then ChokeCasingNo

End Function

Function ChokeCasingYes()
'Created for Pason by Dragon Wood (September 2015)
'Activates the Casing Pressure field if not already active.

    ActiveWorkbook.Worksheets("System Selection").CasingBox.Value = strYes

End Function

Function ChokeCasingNo()
'Created for Pason by Dragon Wood (September 2015)
'Verifies the user does not want a Casing Pressure Sensor

    'Declare the variables
    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExclBox As Integer
        
    'Set the Message Box settings.
    Msg = "Are you sure you do not want to order a Casing Pressure Sensor?"
    Title = "Are You Sure?"
    Config = vbYesNo + vbQuestion + vbDefaultButton2
    ExclBox = MsgBox(Msg, Config, Title)
    
    'Change the monitor selection to Yes.
    If ExclBox = vbNo Then ChokeCasingYes

End Function

Private Sub CAUJBSubJCheck()
'Created for Pason by Dragon Wood (November 2015)
'If a UJB Sub J-Box is ordered, it checks to see if a Sub UJB to UJB Pigtail is ordered.

    If ActiveWorkbook.Worksheets("Choke Actuator").Range("A42").Value = strNo Then Exit Sub

    'Check for a Sub UJB to UJB Pigtail to be ordered.
    If ActiveWorkbook.Worksheets("Choke Actuator").Range("A37").Value = strYes Then Exit Sub
    
    'Warn the user if it wasn't.
    Call CAUJBSubJMessage

End Sub

Function CAUJBSubJMessage()
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
    If ExclBox = vbYes Then Call CAUJBSubJYes

    'Verify the user does not want a Sub UJB to UJB Pigtail.
    If ExclBox = vbNo Then Call CAUJBSubJNo

End Function

Function CAUJBSubJYes()
'Created for Pason by Dragon Wood (November 2015)
'Changes a montior selection to Yes

        With ActiveWorkbook.Worksheets("Choke Actuator")
            .Range("A37").Value = strYes
            .Range("D42").Copy
            .Range("D37").PasteSpecial xlPasteValues
        End With

End Function

Function CAUJBSubJNo()
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
    If ExclBox = vbNo Then CAUJBSubJYes

End Function

Private Sub CASubUJBCableCheck()

    If ActiveWorkbook.Worksheets("Choke Actuator").Range("A42").Value = strNo Then Exit Sub

    'Check for a Sub UJB Sensor Cable to be ordered.
    With ActiveWorkbook.Worksheets("Choke Actuator")
        For Each cell In Range("A31:A35")
            If cell = strYes Then Exit Sub
                Next cell
                Call CASubUJBCableMessage
    End With

End Sub

Function CASubUJBCableMessage()
'Created for Pason by Dragon Wood (November 2015)
'Warns the user they have not ordered a Sub UJB Sensor Cable.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Sub UJB Sensor Cable."
    Title = "Choke Actuator"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("Choke Actuator").Select

End Function

Function ChokeCopyData()
    Dim wkShtMaster As Worksheet
    Dim varSheets As Variant
    Dim lngRow As Long
    Dim lngMatch As Long
    Dim lngIncrement As Long
    Dim shpIcon As Shape
     
    Set wkShtMaster = Worksheets("Master DataList")
         
        With Worksheets("Choke Actuator")
             
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
