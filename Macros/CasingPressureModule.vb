Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Public Sub CasingPressureReset(control As IRibbonControl)
'Created for Pason by Dragon Wood (August 2015)
'Resets the Casing Pressure page so it can be used again.

    Application.ScreenUpdating = False
    
    'Clear all data on the page.
    Call CasingPressureClear
'    Call CasingCopyData
    
    'Close the sheet.
    With ActiveWorkbook.Worksheets("System Selection")
        .CasingBox.Value = strNo
    End With
    
    'Reopen the sheet to refill the required parts.
    With ActiveWorkbook.Worksheets("System Selection")
        .CasingBox.Value = strYes
    End With
    
    Application.ScreenUpdating = True
    Sheets("System Selection").Select

End Sub

Function CasingPressureUnhide()
'Created for Pason by Dragon Wood (August 2015)
'Automatically sets the Required equipment to "Yes" and sets the default value for each. These can of course be changed.

    'Ensure the Casing Pressure sheet is selected.
    With ActiveWorkbook.Worksheets("Casing Pressure")
    
    'Set the Required equipment to Yes.
        .Range("A2:A5").Value = strYes
    
    'Set the default quantity for each item on the list
        .Range("D2:D5").Value = 1
    End With
    
End Function

Function CasingPressureClear()
'Created for Pason by Dragon Wood (August 2015)
'Resets the Casing Pressure Page to the original unused state.

    Application.ScreenUpdating = False
    Application.EnableEvents = False
   
    'Reset the Yes/No field to No on the Casing Pressure Sheet and clear the Quantity Needed Column
    With ActiveWorkbook.Worksheets("Casing Pressure")
        .Visible = True
        .Range("A2:A6").Value = strNo
        .Range("D2:D6").ClearContents
    End With

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Application.GoTo Sheets("Casing Pressure").Range("A1"), True
    Application.GoTo Sheets("System Selection").Range("A1"), True

End Function

Function CasingCopyData()
    Dim wkShtMaster As Worksheet
    Dim varSheets As Variant
    Dim lngRow As Long
    Dim lngMatch As Long
    Dim lngIncrement As Long
    Dim shpIcon As Shape
     
    Set wkShtMaster = Worksheets("Master DataList")
         
        With Worksheets("Casing Pressure")
             
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
