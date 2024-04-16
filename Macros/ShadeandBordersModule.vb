Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Public Sub ShadeSheet()
'Created for Pason by Dragon Wood (August 2015).
'Calls all the shade codes to shade the entire sheet.

    Call ShadeRequired
    Call ShadeChoice
    Call ShadeRecommended
    Call ShadeOptional
    Call ApplyBorders
    
End Sub

Function ShadeRequired()
'Created for Pason by Dragon Wood (August 2015).
'Shades every other row of required parts red and the alternate rows pink.

    'Declare the variables
    Dim lngRow As Long
    Dim lngCol As Long
            
    With ActiveWorkbook.Worksheets("Sheet1")
        For lngCol = 1 To 5
            For lngRow = 2 To 5
                Cells(lngRow, lngCol).Interior.Color = RGB(255, 128, 128)
            Next lngRow
        Next lngCol
    End With
    
End Function

Function ShadeChoice()
'Created for Pason by Dragon Wood (August 2015).
'Shades every other row of required parts, with a choice yellow and the alternate rows dull yellow.

    'Declare the variables
    Dim lngRow As Long
    Dim lngCol As Long
            
    With ActiveWorkbook.Worksheets("Sheet1")
        For lngCol = 1 To 5
            For lngRow = 6 To 10
                Cells(lngRow, lngCol).Interior.Color = RGB(255, 255, 0)
            Next lngRow
        Next lngCol
    End With
    
End Function

Function ShadeRecommended()
'Created for Pason by Dragon Wood (August 2015).
'Shades every other row of recommended parts green and the alternate rows light green.

    'Declare the variables
    Dim lngRow As Long
    Dim lngCol As Long
            
    With ActiveWorkbook.Worksheets("Sheet1")
        For lngCol = 1 To 5
            For lngRow = 11 To 15
                Cells(lngRow, lngCol).Interior.Color = RGB(0, 255, 128)
            Next lngRow
        Next lngCol
    End With
    
End Function

Function ShadeOptional()
'Created for Pason by Dragon Wood (August 2015).
'Shades every other row of optional parts light blue and the alternate rows pale blue.

    'Declare the variables
    Dim lngRow As Long
    Dim lngCol As Long
            
    With ActiveWorkbook.Worksheets("Sheet1")
        For lngCol = 1 To 5
            For lngRow = 16 To 20
                Cells(lngRow, lngCol).Interior.Color = RGB(204, 255, 255)
            Next lngRow
        Next lngCol
    End With
    
End Function

Function ApplyBorders()
'Created for Pason by Dragon Wood (August 2015).
'Applies the borders to the cells on the sheet.

    Range("A2:E20").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With

End Function