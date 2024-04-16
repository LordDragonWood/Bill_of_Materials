Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1

Private Sub Workbook_Open()
'Created for Pason by Dragon Wood (July 2015)

    'Sets the password for each Worksheet, but still allows the code to work.
    Dim wkSheet As Worksheet
    Dim myTime As Variant

    For Each wkSheet In Worksheets
        wkSheet.Protect "P@s0n", UserInterfaceOnly:=True
    Next wkSheet

    'Forces the Workbook to open on the Instructions page.
    Application.GoTo Sheets("Instructions").Range("A1"), True

    myTime = Now() + TimeValue("00:00:05")
        
    Application.OnTime Earliesttime:=myTime, Procedure:="'" & ThisWorkbook.Name & "'!SetSystemSelectionBoxes", Schedule:=True

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
'Created for Pason by Dragon Wood (July 2015)
'Hides all the sheets except the Instructions and saves the workbook

'Unprotect the Workbook
    ActiveWorkbook.Unprotect "P@s0n"
    
'Declare the variables
    Dim GetBook As String
    Dim iDot As Long
    
    GetBook = ThisWorkbook.Name

    iDot = InStrRev(GetBook, ".")
    
    GetBook = Left(GetBook, iDot - 1)

    If GetBook = "Pason Installation BOMs" Then Call HideAll
    
    If Not Me.Saved Then Me.Save

End Sub

Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
'Allows the Yes/No fields to accept partial entries and lower case.

    'Declare the variables
    Dim lngRow As Long
    Dim RngToCheck As Range

    'Set the range to check for change.
    Set RngToCheck = Intersect(sh.Columns(1), Target)

    'No need to do anything if nothing in Column A is changed
    If Not RngToCheck Is Nothing Then

        'Determine the sheets to check
        For Each sht In Array(Sheet3, Sheet4, Sheet5, Sheet6, Sheet7, Sheet8, Sheet9, Sheet10, Sheet11, Sheet12, Sheet13, Sheet17, Sheet18)
            If sht Is sh Then
                Application.ScreenUpdating = False
                Application.EnableEvents = False

      'Determine how many rows there are
      lngRow = Application.Max(2, sh.Range("A" & sh.Rows.Count).End(xlUp).Row) 'if sheet only has headers then the last row number may have been less than 3.

      'Set the acceptance to include lower case and single letters.
      For Each cll In Intersect(Target, sh.Range("A2:A" & lngRow)).Cells
        Select Case cll.Value
          Case "y", "Y", "yes": cll.Value = strYes
          Case "n", "N", "no": cll.Value = strNo
        End Select
      Next cll

      Application.ScreenUpdating = True
        Exit For  'no need to check the rest if one sheet has been found.
        End If
  Next sht
End If

    Application.EnableEvents = True
    Application.EnableEvents = True

End Sub

Private Sub Workbook_WindowActivate(ByVal Wn As Window)
  RemapEnterKey
End Sub

Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)
  UnmapEnterKey
End Sub