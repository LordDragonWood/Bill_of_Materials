Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Sub CreatePDF(control As IRibbonControl)
'Created for Pason by Dragon Wood (July 2015)
'Saves the System Selection & RMS Order page as a PDF file.

  'Declare the variables
    Dim strFile As String
    Dim strCustomer As String
    Dim strRig As String
    Dim strDate As String
        
    strCustomer = ThisWorkbook.Sheets("RMS Order").Range("B2")
    strRig = ThisWorkbook.Sheets("RMS Order").Range("B4")
    strDate = ThisWorkbook.Sheets("System Selection").Range("Z2").Text
    
  'Determines the path to save the PDF file.
    strFile = ThisWorkbook.Path & "\"
    
    'Unhide the RMS Order sheet
    With ActiveWorkbook
    .Worksheets("RMS Order").Visible = True
    End With
          
    'Save the System Selection & RMS Order page as a PDF
    With Sheets("RMS Order")
'        Worksheets("RMS Order").Select
        Worksheets(Array("Rig Survey Form", "RMS Order")).Select
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=strFile & strCustomer & " - " & strRig & " - " & strDate & " - " & "RMS Order" & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    End With
    
End Sub

Sub CreateXLS(control As IRibbonControl)
'Created for Pason by Dragon Wood (July 2015)
'Saves the System Selection & RMS Order pages as a separate Excel Workbook.

    'Declare the variables
    Dim wkbName As Workbook
    Dim wksName As Worksheet
    Dim strFile As String
    Dim strCustomer As String
    Dim strDate As String
    Dim strBook As String
        
    'Assign the variables
    strCustomer = ThisWorkbook.Sheets("System Selection").Range("B6")
    strDate = ThisWorkbook.Sheets("System Selection").Range("Z2").Text
    strBook = ActiveWorkbook.FullName
    strFile = ThisWorkbook.Path & "\"
    
    'Handle Errors
    On Error GoTo EndSub
    Application.DisplayAlerts = False
    
    'Begin the code
    ThisWorkbook.Save
    ThisWorkbook.SaveAs strFile & strCustomer & " - " & strDate & ".xlsx", xlOpenXMLWorkbook
    Set wkbName = ActiveWorkbook
    
    'Unhide the RMS Order sheet
    With ActiveWorkbook
    .Worksheets("RMS Order").Visible = True
    End With
    
    'Delete all the sheets except the System Selection and RMS Order sheets
    For Each wksName In wkbName.Worksheets
        If wksName.Name <> "System Selection" And wksName.Name <> "RMS Order" Then wksName.Visible = False
    Next wksName
    
    'Open the original workbook
    Workbook.Open strBook
    
    'Close the new workbook
    wkbName.Close True
    
    'Error handling code
EndSub:
    Application.DisplayAlerts = True

End Sub
