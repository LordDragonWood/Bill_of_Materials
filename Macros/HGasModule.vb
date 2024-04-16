Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Public Sub HGasReset(control As IRibbonControl)
'Created for Pason by Dragon Wood (August 2015)
'Resets the H-Gas page so it can be used again.

    Application.ScreenUpdating = False
    
    'Clear all data on the page.
    Call HGasClear
'    Call HGasCopyData
    
    'Close the sheet.
    With ActiveWorkbook.Worksheets("System Selection")
        .HGasBox.Value = strNo
    End With
    
    'Reopen the sheet to refill the required parts.
    With ActiveWorkbook.Worksheets("System Selection")
        .HGasBox.Value = strYes
    End With
    
    Application.ScreenUpdating = True
    Sheets("System Selection").Select

End Sub

Function HGasUnhide()
'Created for Pason by Dragon Wood (August 2015)
'Automatically sets the Required equipment to "Yes" and sets the default value for each. These can of course be changed.

    'Ensure the H-Gas sheet is selected.
    With ActiveWorkbook.Worksheets("H-Gas")
    
    'Set the Required equipment to Yes.
        .Range("A2:A4").Value = strYes
    
    'Set the default quantity for each item on the list
        .Range("D2:D3").Value = 1
        .Range("D4").Value = 4
    End With
    
End Function

Function HGasClear()
'Created for Pason by Dragon Wood (August 2015)
'Resets the H-Gas Page to the original unused state.

    Application.ScreenUpdating = False
    Application.EnableEvents = False
   
    'Reset the Yes/No field to No on the H-Gas Sheet and clear the Quantity Needed Column
    With ActiveWorkbook.Worksheets("H-Gas")
        .Visible = True
        .Range("A2:A28").Value = strNo
        .Range("D2:D28").ClearContents
    End With

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Application.GoTo Sheets("H-Gas").Range("A1"), True
    Application.GoTo Sheets("System Selection").Range("A1"), True

End Function

Function HGasCopyData()
    Dim wkShtMaster As Worksheet
    Dim varSheets As Variant
    Dim lngRow As Long
    Dim lngMatch As Long
    Dim lngIncrement As Long
    Dim shpIcon As Shape
     
    Set wkShtMaster = Worksheets("Master DataList")
         
        With Worksheets("H-Gas")
             
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