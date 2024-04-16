Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Sub PreviewOrder(control As IRibbonControl)
'Created for Pason by Dragon Wood (July 2015).
'Unhides the Order Summary sheet.
'Finds the parts and quantities needed and copies them to the Order Summary sheet.

    Application.ScreenUpdating = False
    
    'Declare Variables
    Dim wkSheet As Worksheet
    Dim lngRow As Long
    Dim wksRng As Range
    Dim edrRng As Range
    Dim nSheets
    Dim wsInstance As Range
        
    'Check the Required "Choice" parts.
    If ActiveWorkbook.Worksheets("System Selection").AutoDrillerBox.Value = strYes Then AutoDrillerCheck
    If ActiveWorkbook.Worksheets("System Selection").ChokeBox.Value = strYes Then ChokeActuatorCheck
    If ActiveWorkbook.Worksheets("System Selection").EDRBox.Value = strYes Then EDRCheck
    If ActiveWorkbook.Worksheets("System Selection").ePVTBox.Value = strYes Then ePVTCheck
    If ActiveWorkbook.Worksheets("System Selection").ESRBox.Value = strYes Then ESRCheck
    If ActiveWorkbook.Worksheets("System Selection").PRDBox.Value = strYes Then PRDCheck
    If ActiveWorkbook.Worksheets("System Selection").PVTBox.Value = strYes Then PVTCheck
    If ActiveWorkbook.Worksheets("System Selection").WorkstationsBox.Value = strYes Then WorkstationsCheck
    
    'Unhide the Order Summary sheet.
    Sheets("Order Summary").Visible = True

    'Clear the Order Summary sheet to make sure the order isn't accidentally duplicated.
    Sheets("Order Summary").Range("A2:D2000").ClearContents

    'Declare which sheets to ignore.
    nSheets = Array("Instructions", "System Selection", "Order Summary", "RMS Order", "Master DataList")
        For Each wkSheet In ActiveWorkbook.Worksheets
            If Not IsNumeric(Application.Match(wkSheet.Name, nSheets, 0)) And wkSheet.Visible = True Then
                'Determine how many rows there are for each sheet.
                wkSheet.Select
                'Declare what to do with each sheet not ignored.
                For Each wsInstance In wkSheet.Range("A2:A" & (wkSheet.Range("A" & Rows.Count).End(xlUp).Row))
                    If wsInstance = strYes Then
                        wkSheet.Range(wkSheet.Cells(wsInstance.Row, "B"), wkSheet.Cells(wsInstance.Row, "D")).Copy
                        Sheets("Order Summary").Range("A" & Rows.Count).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
                    End If
                Next wsInstance
            End If
        Next wkSheet
                
Application.ScreenUpdating = True
Application.CutCopyMode = False
Sheets("Order Summary").Select

End Sub

Sub ConsolidateDuplicates(control As IRibbonControl)
'Written for Pason by Dragon Wood (June 2015)
'Combines any duplicate parts and adds the quantities together.

    'Declaring the variables
    Dim lngSum As Long
    Dim lngRowCount As Long
    Dim lngRowField As Long
    Dim lngInterval As Long
    Dim lngTable As Long
    Dim lngCollection As Long
    Dim varPart As Variant
     
    'Ensure the focus is on the Order Summary sheet.
    ActiveWorkbook.Worksheets("Order Summary").Select

    'Count the rows and determine how many there are.
    lngRowCount = Cells(Rows.Count, 1).End(xlUp).Row
     
    Range("B2:B" & lngRowCount).Copy Destination:=Range("Y1")
    ActiveSheet.Range("$Y$1:$Y" & lngRowCount).RemoveDuplicates Columns:=1, Header:=xlNo
     
    lngRowField = Cells(Rows.Count, 25).End(xlUp).Row
    ReDim varPart(1 To lngRowField)
    For lngCollection = 1 To lngRowField
        varPart(lngCollection) = Cells(lngCollection, 25).Value
    Next lngCollection
     
    For lngTable = LBound(varPart) To UBound(varPart)
        lngSum = 0
        For lngInterval = 2 To lngRowCount
            If varPart(lngTable) = Cells(lngInterval, 2) Then
                lngSum = lngSum + Cells(lngInterval, 3).Value
                Cells(lngTable, 24).Value = Cells(lngInterval, 1).Value
            End If
        Next lngInterval
         
        Cells(lngTable, 26).Value = lngSum
    Next lngTable
     
    Range("A2:C" & lngRowCount).Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("X1:Z" & lngRowField).Select
    Selection.Copy
    Range("A2").Select
    Selection.PasteSpecial xlPasteValues
    Range("X1:Z" & lngRowField).ClearContents
    Range("A1").Select
     
End Sub

Public Sub OrderSummaryReset(control As IRibbonControl)
'Created for Pason by Dragon Wood (August 2015)
'Resets the Order Summary page so it can be used again.

    Application.ScreenUpdating = False
    
    'Clear all data on the page.
    Call OrderSummaryClear
    
    'Close the sheet.
    With ActiveWorkbook
        .Worksheets("Order Summary").Visible = False
    End With
    
    Application.ScreenUpdating = True
    
    Sheets("System Selection").Select

End Sub

Function OrderSummaryClear()
'Created for Pason by Dragon Wood (August 2015)
'Resets the Order Summary Page to the original unused state.

    Application.ScreenUpdating = False
       
    'Clears the Order Summary page
    With ActiveWorkbook.Worksheets("Order Summary")
        .Visible = True
        .Range("A2:C2000").ClearContents
    End With

    Application.ScreenUpdating = True
    
    Application.GoTo Sheets("Order Summary").Range("A1"), True
    
    Application.GoTo Sheets("System Selection").Range("A1"), True

End Function