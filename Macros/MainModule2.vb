Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Dim ProgressIndicator As frmProgress
Dim Counter As Long
Dim PctDone As Single
Dim CallMax As Long
Dim CallTotal As Long

Public Sub ResetProject(control As IRibbonControl)
'Created for Pason by Dragon Wood (August 2015)
'Resets the entire workbook to the original unused state.
    
    Call UnhideWorksheets
    Call ResetProjectMessage
    Call ResetPreviewOrderRun
    Call ResetLegend
    Call CloseLegend
    Call HideAll
            
End Sub

Public Sub SystemSelectionReset(control As IRibbonControl)
'Created for Pason by Dragon Wood (August 2015)
'Resets the System Selection page so it can be used again.

    Application.ScreenUpdating = False
    
    'Clear all data on the page.
    Call SystemSelectionClear
        
    Application.ScreenUpdating = True
    Sheets("System Selection").Select

End Sub

Function SystemSelectionClear()
'Created for Pason by Dragon Wood (August 2015)
'Resets the System Selection Page to the original unused state.

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'Reset all the option boxes and clear the fields on the System Selection page
    With ActiveWorkbook.Worksheets("System Selection")
        .Visible = True
        .AutoDrillerBox.Value = strNo
        .CasingBox.Value = strNo
        .ChokeBox.Value = strNo
        .EDRBox.Value = strNo
        .ePVTBox.Value = strNo
        .ESRBox.Value = strNo
        .FlowBox.Value = strYes
        .GABox.Value = strNo
        .HGasBox.Value = strNo
        .PVTBox.Value = strNo
        .PRDBox.Value = strNo
        .ProbeBox.Value = strPrbR
        .SideKickBox.Value = strNo
        .TorqueBox.Value = strTorqueE
        .UJBBox.Value = strYes
        .UnitsBox.Value = strUnitsI
        .VoltageBox.Value = strVolt1
        .WorkstationsBox.Value = strNo
        .Range("B4:B6").ClearContents
        .Range("D4").ClearContents
        .Range("F4").ClearContents
    End With

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    Application.GoTo Sheets("System Selection").Range("A1"), True

End Function

Function ResetAll()
'Created for Pason by Dragon Wood (October 2015)
'Calls all the system clear functions and clears the rest of the data from the workbook.
    
    Application.ScreenUpdating = False
    
    Set ProgressIndicator = New frmProgress
    
    ProgressIndicator.Show vbModeless
    
    Counter = 0
    CallMax = 18
    CallTotal = 5.55
    
    'Call all the Clear Functions for each sheet.
    Call AutoDrillerClear
        Call CloseLegend
        Counter = Counter + 5.55
        PctDone = Counter / (CallMax * CallTotal)
        Call UpdateProgress(PctDone)
    Call CasingPressureClear
        Call CloseLegend
        Counter = Counter + 5.55
        PctDone = Counter / (CallMax * CallTotal)
        Call UpdateProgress(PctDone)
    Call ChokeActuatorClear
        Call CloseLegend
        Counter = Counter + 5.55
        PctDone = Counter / (CallMax * CallTotal)
        Call UpdateProgress(PctDone)
    Call EDRClear
        Call CloseLegend
        Counter = Counter + 5.55
        PctDone = Counter / (CallMax * CallTotal)
        Call UpdateProgress(PctDone)
    Call ePVTClear
        Call CloseLegend
        Counter = Counter + 5.55
        PctDone = Counter / (CallMax * CallTotal)
        Call UpdateProgress(PctDone)
    Call ESRClear
        Call CloseLegend
        Counter = Counter + 5.55
        PctDone = Counter / (CallMax * CallTotal)
        Call UpdateProgress(PctDone)
    Call GasAnalyzerClear
        Call CloseLegend
        Counter = Counter + 5.55
        PctDone = Counter / (CallMax * CallTotal)
        Call UpdateProgress(PctDone)
    Call GeneralUseItemsClear
        Call CloseLegend
        Counter = Counter + 5.55
        PctDone = Counter / (CallMax * CallTotal)
        Call UpdateProgress(PctDone)
    Call HGasClear
        Call CloseLegend
        Counter = Counter + 5.55
        PctDone = Counter / (CallMax * CallTotal)
        Call UpdateProgress(PctDone)
    Call MasterPartsListClear
        Call CloseLegend
        Counter = Counter + 5.55
        PctDone = Counter / (CallMax * CallTotal)
        Call UpdateProgress(PctDone)
    Call OrderSummaryClear
        Call CloseLegend
        Counter = Counter + 5.55
        PctDone = Counter / (CallMax * CallTotal)
        Call UpdateProgress(PctDone)
    Call PRDClear
        Call CloseLegend
        Counter = Counter + 5.55
        PctDone = Counter / (CallMax * CallTotal)
        Call UpdateProgress(PctDone)
    Call PVTClear
        Call CloseLegend
        Counter = Counter + 5.55
        PctDone = Counter / (CallMax * CallTotal)
        Call UpdateProgress(PctDone)
    Call RMSOrderClear
        Call CloseLegend
        Counter = Counter + 5.55
        PctDone = Counter / (CallMax * CallTotal)
        Call UpdateProgress(PctDone)
    Call WorkstationsClear
        Call CloseLegend
        Counter = Counter + 5.55
        PctDone = Counter / (CallMax * CallTotal)
        Call UpdateProgress(PctDone)
    Call RigSurveyFormClear
        Call CloseLegend
        Counter = Counter + 5.55
        PctDone = Counter / (CallMax * CallTotal)
        Call UpdateProgress(PctDone)
    Call SystemSelectionClear
        Call CloseLegend
        Counter = Counter + 5.55
        PctDone = Counter / (CallMax * CallTotal)
        Call UpdateProgress(PctDone)
    Call RSFImportClear
        Call CloseLegend
        Counter = Counter + 5.55
        PctDone = Counter / (CallMax * CallTotal)
        Call UpdateProgress(PctDone)
   
    Unload ProgressIndicator
    Set ProgressIndicator = Nothing
   'Hide the Order Summaryy and RMS Order pages.
    With ActiveWorkbook
    .Worksheets("Order Summary").Visible = False
    .Worksheets("RMS Order").Visible = False
    End With
        
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    Sheets("Instructions").Select

End Function

Sub UpdateProgress(pct)
    With ProgressIndicator
        .FrameProgress.Caption = Format(pct, "0%")
        .LabelProgress.Width = pct * (.FrameProgress.Width - 18)
    End With
    DoEvents
End Sub
Function ResetProjectMessage()
'Created for Pason by Dragon Wood (October 2015)
'Asks the user if they are sure they want to reset the program.

    'Declare the variables
    Dim Msg As String
    Dim Title As String
    Dim Config As Integer
    Dim ExclBox As Integer
        
    'Set the Message Box settings.
    Msg = "This will clear all the data from WorkBook."
    Msg = Msg & vbNewLine & vbNewLine
    Msg = Msg & "Are you sure you're ready to start over?"
    Title = "Are You Sure?"
    Config = vbYesNo + vbExclamation + vbDefaultButton1
    ExclBox = MsgBox(Msg, Config, Title)
    
    'Change the monitor selection to Yes.
    If ExclBox = vbYes Then Call ResetAll
        
End Function

Function HideAll()
'Created for Pason by Dragon Wood (October 2015).
'Hides all the sheets except the Instructions, Rig Survey Form, System Selection, General Use Items, and Mater Parts List pages.

    Application.ScreenUpdating = False

'Hide all the sheets before closing.
    For Each wkSheet In ActiveWorkbook.Worksheets
        If wkSheet.Name <> ("Instructions") Then wkSheet.Visible = xlSheetHidden
    Next wkSheet

'Put the focus of the workbook on the instructions page for the next start up.
    Application.GoTo Sheets("Instructions").Range("A1"), True


End Function

Public Sub RMSOnline(control As IRibbonControl)

    ActiveWorkbook.FollowHyperlink Address:="https://rms.pason.com/pages/Components/ComponentLanding.aspx", NewWindow:=True
    
End Sub

Function VerifyNumeric(Target As Range)
' Checks if Target is empty, otherwise checks if target is numeric. Also removes any spaces. (an added feature)
   
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
'    If Trim(Target) = vbNullString Then Exit Function
     
    If Not IsNumeric(Target.Value) Then
        MsgBox "Enter a number in the quantity field."
        Target.Value = ""
        Application.GoTo Target
         
    ElseIf Target <> Replace(Target, " ", "") Then Target = Replace(Target, " ", "")
    End If
     
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Function

Function ResetPreviewOrderRun()
    
        PreviewOrderRunControl = ActiveWorkbook.Worksheets("Rig Survey Form").Range("AG14") = False

End Function

Public Sub LegendReset(control As IRibbonControl)

    Call ResetLegend

End Sub
 
Function GetShape(ByRef wkSht As Worksheet, ByRef cell As Range) As Shape
    Dim shpIcon As Shape
     
    With wkSht
         
        For Each shpIcon In .Shapes
             
            If shpIcon.TopLeftCell.Row = cell.Row And shpIcon.TopLeftCell.Column = cell.Column Then
                 
                Set GetShape = shpIcon
                Exit Function
            End If
        Next shpIcon
    End With
End Function