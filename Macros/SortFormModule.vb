Rem Attribute VBA_ModuleType=VBAFormModule
Option VBASupport 1
Option Explicit
'Declare Module wide variables.
'These can be used in any routine in this UserForm.
Private nSheets As Variant 'Sheets to ignore.
Private Sub UserForm_Initialize()
'Initialize Module wide Variables

    nSheets = Array("Instructions", "Rig Survey Form", "System Selection", "Order Summary", "RMS Order", "Master DataList", "Master Parts List", "RSFImport")
    
End Sub
Private Sub btnCancel_Click()
'Created for Pason by Dragon Wood (August 2015).
'Closes the Sort List form.

    Unload frmSort

End Sub

Private Sub btnSort_Click()
'Created for Pason by Dragon Wood (August 2015)
'Determines which sort option is selected and calls the correct sub.

    'Declare the variables.
    Dim lngSet As Long
    'Dim lngRng As Range
    
    For lngSet = LBound(nSheets) To UBound(nSheets)
        If ActiveSheet.Name = nSheets(lngSet) Then Exit Sub
    Next lngSet
    
    If Me.Controls("OptionQuantity") Then QuantitySort
    If Me.Controls("OptionPartName") Then PartNameSort
    If Me.Controls("OptionPartNum") Then PartNumSort
    If Me.Controls("OptionOrdered") Then PartOrderedSort
    If Me.Controls("OptionColor") Then ColorSort

End Sub
