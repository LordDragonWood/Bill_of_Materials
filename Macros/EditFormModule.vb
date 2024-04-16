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

    Unload frmEdit
    
End Sub

Private Sub btnEdit_Click()
'Provides a method to add new parts.
'Provides a method to change the priority of the new or existing parts.

    'Declare the variables.
    Dim lngSet As Long
    
    'Make sure they are using a system page.
    For lngSet = LBound(nSheets) To UBound(nSheets)
        If ActiveSheet.Name = nSheets(lngSet) Then Exit Sub
    Next lngSet
    
    'Make sure something is entered in the text box
    If txtEdit.Text = "" Then
        MsgBox "You must enter a Part Number."
        Exit Sub
    End If
    
    With ActiveSheet
        'Make sure they are in the Part Number column
        If Intersect(ActiveCell, .Columns(3)) Is Nothing Then Exit Sub
        
        'Put the Part Number in the field.
        ActiveCell.Value = txtEdit.Text
        ActiveCell.Value = UCase(ActiveCell.Value)
            
        'Set the importance of the part.
        If Me.Controls("optRequired") Then PartRequired
        If Me.Controls("optChoice") Then PartChoice
        If Me.Controls("optRecommended") Then PartRecommended
        If Me.Controls("optOptional") Then PartOptional
    End With
    
    'Clear the controls
    txtEdit.Text = ""
    optRequired = True
    txtEdit.SetFocus
    Unload frmEdit
End Sub

Private Sub PartRequired()
'Sets the part importance to required by coloring the cells appropriately and setting the importance number in Column S.

    Call RadioAction(1)

End Sub

Private Sub PartChoice()
'Sets the part importance to required by coloring the cells appropriately and setting the importance number in Column S.

    Call RadioAction(2)

End Sub

Private Sub PartRecommended()
'Sets the part importance to required by coloring the cells appropriately and setting the importance number in Column S.

    Call RadioAction(3)

End Sub

Private Sub PartOptional()
'Sets the part importance to required by coloring the cells appropriately and setting the importance number in Column S.

    Call RadioAction(4)

End Sub

Private Function RadioAction(ByVal Id As Long) As Boolean
    With ActiveSheet
        .Cells(ActiveCell.Row, "S").Value = Id
        Call ApplyBorders(.Cells(ActiveCell.Row, "A").Resize(, 5))
        Call ApplyColour(.Cells(ActiveCell.Row, "A").Resize(, 5), Id)
    End With
End Function

Private Function ApplyBorders(ByVal rng As Range)
     'Applies the borders to the cells on the sheet.

    With rng

        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone

        .BorderAround LineStyle:=xlContinuous, ColorIndex:=xlAutomatic, Weight:=xlThin

        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
End Function

Private Function ApplyColour(ByVal rng As Range, ByVal Id As Long)
'Applies the color to the cells.

    With rng
    
        Select Case Id
            Case 1
                .Interior.Color = RGB(255, 128, 128)
            Case 2
                .Interior.Color = RGB(255, 255, 0)
            Case 3
                .Interior.Color = RGB(0, 255, 128)
            Case 4
                .Interior.Color = RGB(204, 255, 255)
        End Select
    End With
    
End Function
