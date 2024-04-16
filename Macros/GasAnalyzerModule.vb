Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Public Sub GasAnalyzerReset(control As IRibbonControl)
'Created for Pason by Dragon Wood (August 2015)
'Resets the Gas Analyzer page so it can be used again.

    Application.ScreenUpdating = False
    
    'Clear all data on the page.
    Call GasAnalyzerClear
'    Call GasAnalyzerCopyData
    
    'Close the sheet.
    With ActiveWorkbook.Worksheets("System Selection")
        .GABox.Value = strNo
    End With
    
    'Reopen the sheet to refill the required parts.
    With ActiveWorkbook.Worksheets("System Selection")
        .GABox.Value = strYes
    End With
    
    Application.ScreenUpdating = True
    Sheets("System Selection").Select

End Sub

Function GasAnalyzerUnhide()
'Created for Pason by Dragon Wood (August 2015)
'Automatically sets the Required equipment to "Yes" and sets the default value for each. These can of course be changed.

    'Ensure the Gas Analyzer sheet is selected.
    With ActiveWorkbook.Worksheets("Gas Analyzer")
    
    'Set the Required equipment to Yes.
        .Range("A2:A4").Value = strYes
    
    'Set the default quantity for each item on the list
        .Range("D2:D4").Value = 1
    End With
    
End Function

Function GasAnalyzerClear()
'Created for Pason by Dragon Wood (August 2015)
'Resets the Gas Analyzer Page to the original unused state.

    Application.ScreenUpdating = False
    Application.EnableEvents = False
   
    'Reset the Yes/No field to No on the Gas Analyzer Sheet and clear the Quantity Needed Column
    With ActiveWorkbook.Worksheets("Gas Analyzer")
        .Visible = True
        .Range("A2:A30").Value = strNo
        .Range("D2:D30").ClearContents
    End With

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Application.GoTo Sheets("Gas Analyzer").Range("A1"), True
    Application.GoTo Sheets("System Selection").Range("A1"), True

End Function

Function GasAnalyzerCopyData()
    Dim wkShtMaster As Worksheet
    Dim varSheets As Variant
    Dim lngRow As Long
    Dim lngMatch As Long
    Dim lngIncrement As Long
    Dim shpIcon As Shape
     
    Set wkShtMaster = Worksheets("Master DataList")
         
        With Worksheets("Gas Analyzer")
             
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
