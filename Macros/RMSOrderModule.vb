Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Sub RMSOrderFill(control As IRibbonControl)
'Created for Pason by Dragon Wood (July 2015).
'Copies the final order quantities from the Order Summary page to the RMS Order page.
'Then sorts the data so it's in alphabetical order by Part Number

Application.ScreenUpdating = False
    
    'Declare the variables.
    Dim lngRow As Long

    'Unhide the RMS Order sheet.
    With ActiveWorkbook
    .Worksheets("RMS Order").Visible = True
    End With
    
    'Clear the RMS Order sheet to make sure the order isn't accidentally duplicated
    Sheets("RMS Order").Range("A13:D2000").ClearContents

    'Determine how many rows in the Order Summary sheet need copied.
    With ActiveWorkbook
    .Worksheets("Order Summary").Select
    
    lngRow = Range("A" & Rows.Count).End(xlUp).Row
    
    'Copy each rows data.
    For Each cell In Range("A2:A" & lngRow)
        Range(Cells(cell.Row, "A"), Cells(cell.Row, "A")).Copy
        Sheets("RMS Order").Range("A" & Rows.Count).End(xlUp).Offset(12, 0).PasteSpecial xlPasteValues
        Range(Cells(cell.Row, "B"), Cells(cell.Row, "B")).Copy
        Sheets("RMS Order").Range("B" & Rows.Count).End(xlUp).Offset(12, 0).PasteSpecial xlPasteValues
        Range(Cells(cell.Row, "C"), Cells(cell.Row, "C")).Copy
        Sheets("RMS Order").Range("C" & Rows.Count).End(xlUp).Offset(12, 0).PasteSpecial xlPasteValues
        
Next cell

    End With
    
    Call RMSFillHeader
    Call RMSSort
    Call RMSCenter
    Call RemoveReapplyBorders
    Call RMSNotes

    
Application.ScreenUpdating = True
Application.CutCopyMode = False

'Reset the RMS Order sheet as the focus.
Application.GoTo Sheets("RMS Order").Range("A1"), True

End Sub

Public Sub RMSOrderReset(control As IRibbonControl)
'Created for Pason by Dragon Wood (August 2015)
'Resets the RMS Order page so it can be used again.

    Application.ScreenUpdating = False
    
    'Clear all data on the page.
    Call RMSOrderClear
    
    'Close the sheet.
    With ActiveWorkbook
        .Worksheets("RMS Order").Visible = False
    End With

    Application.ScreenUpdating = True
    Sheets("System Selection").Select

End Sub

Function RMSOrderClear()
'Created for Pason by Dragon Wood (August 2015)
'Resets the RMS Order Page to the original unused state.

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    With ActiveWorkbook.Worksheets("RMS Order")
        .Visible = True
        .Range("A13:D" & Rows.Count).Clear
        .Range("B2:C10").ClearContents
    End With

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Application.GoTo Sheets("RMS Order").Range("A1"), True
    Application.GoTo Sheets("System Selection").Range("A1"), True

End Function

Function RemoveReapplyBorders()
'Created for Pason by Dragon Wood (August 2015)
    
    Dim lngRows As Long
    Dim lastrow As Long
     
    With Worksheets("RMS Order")
        lastrow = .Cells(.Rows.Count, 1).End(xlUp).Row
        If lastrow = 12 Then Exit Function
        For lngRows = 13 To lastrow
            .Range("A" & lngRows & ":D" & lngRows).Borders(xlEdgeBottom).LineStyle = xlContinuous
        Next
    End With

End Function

Function RMSFillHeader()
'Created for Pason by Dragon Wood (September 2015).

    Dim piCustomer As String
    Dim piRigName As String
    Dim piLocation As String
    Dim piFieldTech As String
    Dim piEstShipDate As String
    
    piCustomer = ThisWorkbook.Sheets("System Selection").Range("B4")
    piRigName = ThisWorkbook.Sheets("Rig Survey Form").Range("N4")
    piLocation = ThisWorkbook.Sheets("Rig Survey Form").Range("M6")
    piFieldTech = ThisWorkbook.Sheets("System Selection").Range("B6")
    piEstShipDate = ThisWorkbook.Sheets("System Selection").Range("Z2").Text

    With ActiveWorkbook.Worksheets("RMS Order")
        .Range("B2:C2").Value = piCustomer
        .Range("B4:C4").Value = piRigName
        .Range("B6:C6").Value = piLocation
        .Range("B8:C8").Value = piFieldTech
        .Range("B10:C10").Value = piEstShipDate
    End With
    
End Function

Function RMSSort()
'Created for Pason by Dragon Wood (September 2015)

    Dim lngRng As Range

    'Ensure the RMS Order page is selected.
    With ActiveWorkbook.Worksheets("RMS Order")

        'Determine how many rows need sorted.
        Set lngRng = .UsedRange.Offset(12).Resize(.UsedRange.Rows.Count - 1)

        'Apply the sort.
        With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=lngRng.Columns("B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange lngRng
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With

End Function

Function RMSCenter()
'Created for Pason by Dragon Wood (September 2015)
    
    Dim lngRows As Long
    Dim lastrow As Long
     
    With Worksheets("RMS Order")
        lastrow = .Cells(.Rows.Count, 1).End(xlUp).Row
        If lastrow = 12 Then Exit Function
        For lngRows = 13 To lastrow
            .Range("C" & lngRows).HorizontalAlignment = xlCenter
            .Range("C" & lngRows).VerticalAlignment = xlCenter
        Next
    End With

End Function

Function RMSNotes()
'Created for Pason by Dragon Wood (November 2015)
    
    Dim lngRows As Long
    Dim lastrow As Long
     
    With Worksheets("RMS Order")
        lastrow = .Cells(.Rows.Count, 1).End(xlUp).Row
        If lastrow = 12 Then Exit Function
        For lngRows = 13 To lastrow
            .Range("D" & lngRows).Locked = False
        Next
    End With

End Function
