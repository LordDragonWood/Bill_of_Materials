Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Public Sub AutoDrillerReset(control As IRibbonControl)
'Created for Pason by Dragon Wood (August 2015)
'Resets the AutoDriller page so it can be used again.

    Application.ScreenUpdating = False
    
    'Clear all data on the page.
    Call AutoDrillerClear
    
    'Close the sheet.
    With ActiveWorkbook.Worksheets("System Selection")
        .AutoDrillerBox.Value = strNo
    End With
    
    'Reopen the sheet to refill the required parts.
    With ActiveWorkbook.Worksheets("System Selection")
        .AutoDrillerBox.Value = strYes
    End With
    
    Application.ScreenUpdating = True
    
    Sheets("System Selection").Select


End Sub

Function AutoDrillerUnhide()
'Created for Pason by Dragon Wood (August 2015)
'Automatically sets the Required equipment to "Yes" and sets the default value for each. These can of course be changed.

    'Ensure the AutoDriller sheet is selected.
    With ActiveWorkbook.Worksheets("AutoDriller")
    
    'Set the Required equipment to Yes.
        .Range("A2:A11").Value = strYes
    
    'Set the default quantity for each item on the list
        .Range("D2:D7").Value = 1
        .Range("D8").Value = 2
        .Range("D9:D11").Value = 1
    End With
    
End Function

Function AutoDrillerClear()
'Created for Pason by Dragon Wood (August 2015)
'Resets the AutoDriller Page to the original unused state.

    Application.ScreenUpdating = False
    Application.EnableEvents = False
   
    'Reset the Yes/No field to No on the AutoDriller Sheet and clear the Quantity Needed Column
    With ActiveWorkbook.Worksheets("AutoDriller")
        .Visible = True
        .Range("A2:A50").Value = strNo
        .Range("D2:D50").ClearContents
    End With
 
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Application.GoTo Sheets("AutoDriller").Range("A1"), True
    Application.GoTo Sheets("System Selection").Range("A1"), True

End Function

Public Sub AutoDrillerCheck()
'Created for Pason by Dragon Wood (August 2015)
'Checks to make sure that the items that are required, but need a choice are chosen.
    
    Call StepperMotorCableCheck
    Call EncoderCableCheck
    Call CrimpSleevesCheck
    Call EncoderCheck
    Call ADRUJBSubJCheck
    Call ADSubUJBCableCheck

End Sub

Private Sub StepperMotorCableCheck()

    'Check for a Stepper Motor Cable to be ordered.
    With ActiveWorkbook.Worksheets("AutoDriller")
        For Each cell In Range("A12:A13")
            If cell = strYes Then Exit Sub
                Next cell
                Call StepperMotorCableMessage
    End With

End Sub

Function StepperMotorCableMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a Stepper Motor Cable.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer
    
    Msg = "You did not order a Stepper Motor Cable."
    Title = "AutoDriller"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("AutoDriller").Select
    
End Function

Private Sub EncoderCableCheck()
    
    'Check for an Encoder Cable to be ordered.
    With ActiveWorkbook.Worksheets("AutoDriller")
        For Each cell In Range("A14:A15")
            If cell = strYes Then Exit Sub
                Next cell
                Call EncoderCableMessage
    End With

End Sub

Function EncoderCableMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered an Encoder Cable.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer
    
    Msg = "You did not order an Encoder Cable."
    Title = "AutoDriller"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("AutoDriller").Select
    
End Function

Private Sub CrimpSleevesCheck()

        'Check for a Cable Crimps to be ordered.
    With ActiveWorkbook.Worksheets("AutoDriller")
        For Each cell In Range("A16:A17")
            If cell = strYes Then Exit Sub
                Next cell
                Call CrimpSleevesMessage
    End With

End Sub

Function CrimpSleevesMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered Crimp Sleeves.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer
    
    Msg = "You did not order Crimp Sleeves."
    Title = "AutoDriller"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("AutoDriller").Select
    
End Function

Private Sub EncoderCheck()

    'Check for an Encoder Cable to be ordered.
    With ActiveWorkbook.Worksheets("AutoDriller")
        For Each cell In Range("A18:A19")
            If cell = strYes Then Exit Sub
                Next cell
                Call EncoderMessage
    End With

End Sub

Function EncoderMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered an Encoder.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer
    
    Msg = "You did not order an Encoder."
    Title = "AutoDriller"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("AutoDriller").Select
    
End Function

Private Sub ADRUJBSubJCheck()
'Created for Pason by Dragon Wood (November 2015)
'If a UJB Sub J-Box is ordered, it checks to see if a Sub UJB to UJB Pigtail is ordered.

    If ActiveWorkbook.Worksheets("AutoDriller").Range("A50").Value = strNo Then Exit Sub

    'Check for a Sub UJB to UJB Pigtail to be ordered.
    If ActiveWorkbook.Worksheets("AutoDriller").Range("A47").Value = strYes Then Exit Sub
    
    'Warn the user if it wasn't.
    Call ADRUJBSubJMessage
    
End Sub

Function ADRUJBSubJMessage()
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
    If ExclBox = vbYes Then Call ADRUJBSubJYes

    'Verify the user does not want a Sub UJB to UJB Pigtail.
    If ExclBox = vbNo Then Call ADRUJBSubJNo

End Function

Function ADRUJBSubJYes()
'Created for Pason by Dragon Wood (November 2015)
'Changes a montior selection to Yes

        With ActiveWorkbook.Worksheets("AutoDriller")
            .Range("A47").Value = strYes
            .Range("D50").Copy
            .Range("D47").PasteSpecial xlPasteValues
        End With

End Function

Function ADRUJBSubJNo()
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
    If ExclBox = vbNo Then ADRUJBSubJYes

End Function

Private Sub ADSubUJBCableCheck()

    If ActiveWorkbook.Worksheets("AutoDriller").Range("A50").Value = strNo Then Exit Sub
    
    'Check for a Sub UJB Sensor Cable to be ordered.
    With ActiveWorkbook.Worksheets("AutoDriller")
        For Each cell In Range("A41:A45")
            If cell = strYes Then Exit Sub
                Next cell
                Call ADSubUJBCableMessage
    End With

End Sub

Function ADSubUJBCableMessage()
'Created for Pason by Dragon Wood (November 2015)
'Warns the user they have not ordered a Sub UJB Sensor Cable.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a Sub UJB Sensor Cable."
    Title = "AutoDriller"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("AutoDriller").Select

End Function

Function AutoDrillerCopyData()
    Dim wkShtMaster As Worksheet
    Dim varSheets As Variant
    Dim lngRow As Long
    Dim lngMatch As Long
    Dim lngIncrement As Long
    Dim shpIcon As Shape
     
    Set wkShtMaster = Worksheets("Master DataList")
         
        With Worksheets("AutoDriller")
             
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
