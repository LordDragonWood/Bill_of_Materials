Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Public Sub GeneralUseItemsReset(control As IRibbonControl)
'Created for Pason by Dragon Wood (August 2015)
'Resets the General Use Items page so it can be used again.

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'Clear all data on the page.
    Call GeneralUseItemsClear
    Call GeneralUseItemsCopyData
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Function GeneralUseItemsClear()
'Created for Pason by Dragon Wood (August 2015)
'Resets the General Use Items Page to the original unused state.

    Application.ScreenUpdating = False
    Application.EnableEvents = False
   
    'Reset the Yes/No field to No on the General Use Items Sheet and clear the Quantity Needed Column
    With ActiveWorkbook.Worksheets("General Use Items")
        .Range("A2:A39").Value = strNo
        .Range("D2:D39").ClearContents
    End With

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Application.GoTo Sheets("General Use Items").Range("A1"), True
    
End Function

Function GeneralUseItemsCopyData()
    Dim wkShtMaster As Worksheet
    Dim varSheets As Variant
    Dim lngRow As Long
    Dim lngMatch As Long
    Dim lngIncrement As Long
    Dim shpIcon As Shape
     
    Set wkShtMaster = Worksheets("Master DataList")
         
        With Worksheets("General Use Items")
             
            lngRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            For lngIncrement = 2 To lngRow
                 
                lngMatch = 0
                On Error Resume Next
                lngMatch = Application.Match(.Cells(lngIncrement, "C"), wkShtMaster.Columns("A"), 0)
                On Error GoTo 0
                If lngMatch > 0 Then
                     
                    .Cells(lngIncrement, "B").Value = wkShtMaster.Cells(lngMatch, "B").Value
                    Set shpIcon = GetShape(wkShtMaster, .Cells(lngMatch, "C"))
                    If Not shpIcon Is Nothing Then
                        .Hyperlinks.Add Anchor:=.Cells(lngIncrement, "E"), Address:=shpIcon.Hyperlink.Address, SubAddress:="", TextToDisplay:="Image"
                    End If
                End If
            Next lngIncrement
             
        End With
End Function