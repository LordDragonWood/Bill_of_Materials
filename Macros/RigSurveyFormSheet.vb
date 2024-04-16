Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    With ActiveWorkbook.Worksheets("Rig Survey Form")
        .Range("AG7").Value = Range("D4").Value
        .Range("AG8").Value = Range("C6").Value
    End With

End Sub
