Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Public Sub PRDReset(control As IRibbonControl)
'Created for Pason by Dragon Wood (August 2015)
'Resets the PRD page so it can be used again.

    Application.ScreenUpdating = False
    
    'Clear all data on the page.
    Call PRDClear
'    Call PRDCopyData
    
    'Close the sheet.
    With ActiveWorkbook.Worksheets("System Selection")
        .PRDBox.Value = strNo
    End With
    
    'Reopen the sheet to refill the required parts.
    With ActiveWorkbook.Worksheets("System Selection")
        .PRDBox.Value = strYes
    End With
    
    Application.ScreenUpdating = True
    Sheets("System Selection").Select

End Sub

Function PRDUnhide()
'Created for Pason by Dragon Wood (August 2015)
'Automatically sets the Required equipment to "Yes" and sets the default value for each. These can of course be changed.

    'Ensure the PRD sheet is selected.
    With ActiveWorkbook.Worksheets("PRD")
    
        'Set the Required equipment to Yes.
        .Range("A2:A3").Value = strYes
    
        'Set the default quantity for each item on the list
        .Range("D2:D3").Value = 1
    End With
    
End Function

Function PRDClear()
'Created for Pason by Dragon Wood (August 2015)
'Resets the PRD Page to the original unused state.

    Application.ScreenUpdating = False
    Application.EnableEvents = False
   
    'Reset the Yes/No field to No on the PRD Sheet and clear the Quantity Needed Column
    With ActiveWorkbook.Worksheets("PRD")
        .Visible = True
        .Range("A2:A25").Value = strNo
        .Range("D2:D25").ClearContents
    End With

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Application.GoTo Sheets("PRD").Range("A1"), True
    Application.GoTo Sheets("System Selection").Range("A1"), True

End Function

Function PRDUJB()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they cannot install an PRD without a UJB and asks if they want to add a UJB to the order.

    'Declare the variables
    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExclBox As Integer
        
    'Set the Message Box settings.
    Msg = "You cannot install a Rig Display without a UJB."
    Msg = Msg & vbNewLine & vbNewLine
    Msg = Msg & "Would you like to add the UJB to your order?"
    Title = "Warning!"
    Config = vbYesNo + vbExclamation + vbDefaultButton1
    ExclBox = MsgBox(Msg, Config, Title)
    
    'Activate the UJB box if Yes.
    If ExclBox = vbYes Then UJBYes
    
    'Close the PRD sheet if No.
    If ExclBox = vbNo Then PRDUJBNo

End Function

Function PRDUJBNo()
'Created for Pason by Dragon Wood (August 2015)
'Activates the UJB field if not already active

    ActiveWorkbook.Worksheets("System Selection").PRDBox.Value = strNo
    ActiveWorkbook.Worksheets("System Selection").Select

End Function

Public Sub PRDCheck()
'Created for Pason by Dragon Wood (August 2015)
'Checks to make sure that the items that are required, but need a choice are chosen.

    Call PRDMountCheck

End Sub

Private Sub PRDMountCheck()

    'Check for a PRD Mount to be ordered.
    With ActiveWorkbook.Worksheets("PRD")
        For Each cell In Range("A4:A6")
            If cell = strYes Then Exit Sub
                Next cell
                Call PRDMountMessage
    End With

End Sub

Function PRDMountMessage()
'Created for Pason by Dragon Wood (August 2015)
'Warns the user they have not ordered a PRD Mount.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You did not order a mount for the PRD."
    Title = "Rig Display"
    Config = vbOKOnly + vbInformation
    ExcelBox = MsgBox(Msg, Config, Title)
    Sheets("PRD").Select

End Function

Function PRDCopyData()
    Dim wkShtMaster As Worksheet
    Dim varSheets As Variant
    Dim lngRow As Long
    Dim lngMatch As Long
    Dim lngIncrement As Long
    Dim shpIcon As Shape
     
    Set wkShtMaster = Worksheets("Master DataList")
         
        With Worksheets("PRD")
             
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
