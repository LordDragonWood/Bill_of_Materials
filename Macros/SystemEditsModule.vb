Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Sub EditSelection(control As IRibbonControl)

    frmEdit.Show
        
End Sub

Sub RemoveSelection(control As IRibbonControl)
    
    With ActiveSheet
        If Intersect(ActiveCell, .Columns(3)) Is Nothing Then Exit Sub
        If Intersect(ActiveCell, .UsedRange) Is Nothing Then Exit Sub
        
        Call RemoveSelectionMessage
    
    End With
        
End Sub

Function RemoveSelectionMessage()
'Warns the user they are about to permanently delete something.

    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExcelBox As Integer

    Msg = "You are about to permanently delete this item from the list."
    Msg = Msg & vbNewLine & vbNewLine
    Msg = Msg & "Are you sure you want to continue?"
    Title = "WARNING!!"
    Config = vbYesNo + vbExclamation + vbDefaultButton1
    ExcelBox = MsgBox(Msg, Config, Title)
    
    'Remove the part if selection is Yes.
    If ExcelBox = vbYes Then Call RemoveSelectionYes
    
    'Cancel the proceedure if the selection is No.
    If ExcelBox = vbNo Then Exit Function

End Function

Function RemoveSelectionYes()
'Removes the entire row for the selected item.

    With ActiveSheet
        .Unprotect "P@s0n"
        ActiveCell.EntireRow.Delete
        .Protect "P@s0n"
    End With
        
End Function