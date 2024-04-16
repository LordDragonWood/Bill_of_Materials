Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Public Sub RigSurveyFormPopulate(control As IRibbonControl)
'Created for Pason by Dragon Wood (October 2015)
'Imports the data from the Customer filled Rig Survey Form.
    
    Dim wbMaster As Workbook, wbCustomer  As Workbook
    Dim wsImport As Worksheet, ws As Worksheet
    Dim aCustomers As Variant
    Dim iCustomer As Long
     
    Call UnhideWorksheets
    Call RSFImportClear
    Call RigSurveyFormClear
    
    Set wbMaster = ThisWorkbook
    Set wsImport = wbMaster.Worksheets("RSFImport")
     
     'ask for customer wb name or names
    aCustomers = Application.GetOpenFilename("*.xls?, Customer Files", , "Select Customer Workbook(s)", , True)
     
    If Not IsArray(aCustomers) Then Exit Sub
     
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    For iCustomer = LBound(aCustomers) To UBound(aCustomers)
         
         'open each wb
        Workbooks.Open Filename:=aCustomers(iCustomer)
        Set wbCustomer = ActiveWorkbook
         
         'look for special marker on each sheet (one / WB)
        For Each ws In wbCustomer.Worksheets
            If ws.Cells(1, 1).Value = "Contractor" Then
                ws.Cells(1, 1).CurrentRegion.Columns(2).Copy
                wbMaster.Activate
                wsImport.Visible = True
                wsImport.Select
                wsImport.Cells(wsImport.Rows.Count, 1).End(xlUp).Offset(1, 0).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
                wsImport.Visible = False
                
                Exit For
            End If
             
        Next
         
        wbCustomer.Close (False)
         
    Next iCustomer
    
    Application.GoTo Sheets("Rig Survey Form").Range("A1")
    
    Call RigSurveyFormFill
    Call RigSurveyFormCleanUp
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
'    Application.GoTo Sheets("Rig Survey Form").Range("A1")
    
End Sub

Public Sub RigSurveyFormReset(control As IRibbonControl)
'Created for Pason by Dragon Wood (September 2015)
'Resets the Rig Survey Form page so it can be used again.

    Application.ScreenUpdating = False
    
    'Clear all data on the page.
    Call RigSurveyFormClear
    Call RSFImportClear
    
    Application.ScreenUpdating = True
    Sheets("Rig Survey Form").Range("D4").Select

End Sub

Function RigSurveyFormClear()
'Created for Pason by Dragon Wood (November 2015).
'Clears all the form fields that are not actually used.

    Dim CellsToCheck As Variant
    Dim cll As Variant

    Application.ScreenUpdating = False

    'Determine how many rows in the Order Summary sheet need copied.
    With ActiveWorkbook.Worksheets("Rig Survey Form")
    
    CellsToCheck = Array("D4:I4", "N4:S4", "C6:I6", "M6:S6", "C8:I8", "M8:S8", "C12:E12", "J12:K12", "R12:S12", "C14:D14", "I14:J14", "N14:P14", "C16:E16", "K16:M16", "R16:S16", "C18:D18", "I18:J18", "Q18:R18", "D20:H20", "Q20:R20", "D24:F24", "L24:N24", "E26:F26", "L26:N26", "E28:F28", "L28:N28", "E30:F30", "D34:T34", "A36:T36", "D38:T38", "A40:T40", "C44:T44", "A46:T46", "E48:F48", "K48:L48", "C50:D50", "H50:I50", "M50:N50", "R50:S50", "C52:D52", "H52:I52", "M52:N52", "R52:S52", "C54:D54", "H54:I54", "M54:N54", "R54:S54", "C56:D56", "H56:I56", "M56:N56", "R56:S56", "C58:D58", "H58:I58", "M58:N58", "R58:S58", "C60:D60", "H60:I60", "M60:N60", "R60:S60", "C62:D62", "H62:I62", "M62:N62", "R62:S62", "C64:D64", "H64:I64", "M64:N64", "R64:S64", "J66:T66", "A68:T68")

    For Each cll In CellsToCheck
        With .Range(cll)
            .ClearContents
        End With

    Next cll

    End With

    Application.ScreenUpdating = True

End Function

Function RSFImportClear()
'Created for Pason by Dragon Wood (October 2015)
'Clears the Rig Survey Form Import page.

    Application.ScreenUpdating = False

    With ActiveWorkbook.Worksheets("RSFImport")
        .Range("A2:BQ2").ClearContents
    End With
    
    Application.ScreenUpdating = True

End Function

Function RigSurveyFormFill()
'Created for Pason by Dragon Wood (October 2015)
'Fills in the Rig Survey Form with the data imported from the Customer version.

    Application.ScreenUpdating = False
    
    'Declare the variables.
    
    'Unhide the RSFImport sheet.
    ActiveWorkbook.Worksheets("RSFImport").Visible = True
    
    'Determine how many rows in the Order Summary sheet need copied.
    With ActiveWorkbook
    .Worksheets("RSFImport").Select
    
    DestnCellsAddresses = Array("D4", "N4", "C6", "M6", "C8", "M8", "C12", "J12", "R12", "C14", "I14", "N14", "C16", "K16", "R16", "C18", "I18", "Q18", "D20", "Q20", "D24", "L24", "E26", "L26", "E28", "L28", "E30", "D34", "A36", "D38", "A40", "C44", "A46", "E48", "K48", "C50", "H50", "M50", "R50", "C52", "H52", "M52", "R52", "C54", "H54", "M54", "R54", "C56", "H56", "M56", "R56", "C58", "H58", "M58", "R58", "C60", "H60", "M60", "R60", "C62", "H62", "M62", "R62", "C64", "H64", "M64", "R64", "J66", "A68")

    i = 0

    For Each cll In Range("A2:BQ2").Cells
    Sheets("Rig Survey Form").Range(DestnCellsAddresses(i)).Value = cll.Value
    i = i + 1

    Next cll

    End With
    
    ActiveWorkbook.Worksheets("RSFImport").Visible = False
    Application.GoTo Sheets("Rig Survey Form").Range("A1")
    Application.ScreenUpdating = True
    
End Function

Function RigSurveyFormCleanUp()
'Created for Pason by Dragon Wood (November 2015).
'Clears all the form fields that are not actually used.

    Dim CellsToCheck As Variant
    Dim CheckedCells As Variant
    Dim CellRanges As Range

    Application.ScreenUpdating = False

    'Determine how many rows in the Order Summary sheet need copied.
    With ActiveWorkbook.Worksheets("Rig Survey Form")
    
        CellsToCheck = Array("D4:I4", "N4:S4", "C6:I6", "M6:S6", "C8:I8", "M8:S8", "C12:E12", "J12:K12", "R12:S12", "C14:D14", "I14:J14", "N14:P14", "C16:E16", "K16:M16", "R16:S16", "C18:D18", "I18:J18", "Q18:R18", "D20:H20", "Q20:R20", "D24:F24", "L24:N24", "E26:F26", "L26:N26", "E28:F28", "L28:N28", "E30:F30", "D34:T34", "A36:T36", "D38:T38", "A40:T40", "C44:T44", "A46:T46", "E48:F48", "K48:L48", "C50:D50", "H50:I50", "M50:N50", "R50:S50", "C52:D52", "H52:I52", "M52:N52", "R52:S52", "C54:D54", "H54:I54", "M54:N54", "R54:S54", "C56:D56", "H56:I56", "M56:N56", "R56:S56", "C58:D58", "H58:I58", "M58:N58", "R58:S58", "C60:D60", "H60:I60", "M60:N60", "R60:S60", "C62:D62", "H62:I62", "M62:N62", "R62:S62", "C64:D64", "H64:I64", "M64:N64", "R64:S64", "J66:T66", "A68:T68")

        For Each CheckedCells In CellsToCheck
            With .Range(CheckedCells)
                If .MergeCells Then
                    If .Cells(1, 1).Value = 0 Then .ClearContents
                Else
                    For Each CellRanges In .Cells
                        If CellRanges.Value = 0 Then CellRanges.ClearContents
                    Next
                End If
            End With
        Next
    End With

    Application.ScreenUpdating = True
    
End Function
