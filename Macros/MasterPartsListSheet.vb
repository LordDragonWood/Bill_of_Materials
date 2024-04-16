Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1

Private Sub Worksheet_Change(ByVal Target As Range)

    'Absolutely prevent this sub if EnableEvents is False. Probably Redundant
    If Application.EnableEvents = False Then Exit Sub
     
'    If Not Intersect(Target, Range("A3:A1958")) Is Nothing Then VerifyProperCaseYesNo Target
'    If Not Intersect(Target, Range("D3:D1958")) Is Nothing Then VerifyNumeric Target

End Sub