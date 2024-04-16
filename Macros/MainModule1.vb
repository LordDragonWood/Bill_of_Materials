Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
'Decalre the Constants and Public Variables for the entire workbook.
Public Const strYes As String = "Yes"
Public Const strNo As String = "No"
Public Const strVolt1 As String = "110V"
Public Const strVolt2 As String = "220V"
Public Const strPrbR As String = "Radar"
Public Const strPrbM As String = "Mud Probe"
Public Const strPrbB As String = "Both"
Public Const strUnitsI As String = "Imperial"
Public Const strUnitsM As String = "Metric"
Public Const strTorqueE As String = "Electric"
Public Const strTorqueH As String = "Hydraulic"
Public Const ADRPowerBarPN1 As String = "ADR005"
Public Const ADRPowerBarPN2 As String = "PWRASS006"
Public Const SurgeSupressorPN1 As String = "PWR021"
Public Const SurgeSupressorPN2 As String = "PWR084"
Public Const CMPPowerCordPN1 As String = "CBL035"
Public Const CMPPowerCordPN2 As String = "CBL076"
Public Const TorqueEPN1 As String = "SEN008"
Public Const TorqueEPN2 As String = "SENASS109"

Public TorqueHPN As String

Public PreviewOrderRunControl As Boolean

Public Sub ActivateWorkbook(control As IRibbonControl)
'Created for Pason by Dragon Wood (July 2015)
'Activates workbook so it can be used.

    Call UnhideWorksheets
    
    Application.GoTo Sheets("System Selection").Range("A1"), True
    
End Sub

Function UJBYes()
'Created for Pason by Dragon Wood (August 2015)
'Activates the UJB field if not already active

    ActiveWorkbook.Worksheets("System Selection").UJBBox.Value = strYes

End Function

Public Sub SaveProject(control As IRibbonControl)
'Created for Pason by Dragon Wood (October 2015).
    Dim GetBook As String
    Dim iDot As Long
    
    GetBook = ThisWorkbook.Name

    iDot = InStrRev(GetBook, ".")
    
    GetBook = Left(GetBook, iDot - 1)

    If GetBook = "Pason Installation BOMs" Then
        Call SaveProjectAs
    Else
        Call SaveWorkbook
    End If

End Sub
Public Sub NewProject(control As IRibbonControl)

    Call SaveProjectAs

End Sub

Function SaveProjectAs()
'Created for Pason by Dragon Wood (October 2015)
'Saves the file as a new file to preserve the data in the project.

'Declare the Variables for Saving the File
    Dim fileSaveName As String
    Dim rigSaveName As String
    Dim customerSaveName As String

'Declare the Variables for the Directory Path
    Dim fileRootPath As String
    Dim fileSavePath As String

'Declare the Varialbles for the Input Boxes
    Dim contractorInput As String
    Dim rigInput As String
    Dim operatorInput As String
    Dim customerInput As String

 'Unhide the sheets if still hidden
    Call UnhideWorksheets

 'Make the System Selection page the focus point
    Application.GoTo Sheets("Rig Survey Form").Range("A1"), True

 'Check the Customer Name & Field Tech Name fields for content. If there, use the content, if not provide an input box for entering the data.

    With Sheets("Rig Survey Form")
        If .Range("D4").Value = "" Then
            .Range("D4").Select
            contractorInput = InputBox("Please enter the Drilling Contractor Name.", "Drilling Contractor Name")
            ActiveCell.FormulaR1C1 = contractorInput
        End If
        If .Range("N4").Value = "" Then
            .Range("N4").Select
            rigInput = InputBox("Please fill in the Rig Name & Number Field", "Rig Name & Number")
            ActiveCell.FormulaR1C1 = rigInput
        End If
        If .Range("C6").Value = "" Then
            .Range("C6").Select
            operatorInput = InputBox("Please fill in the Operator Name.", "Operator Name")
            ActiveCell.FormulaR1C1 = operatorInput
        End If
    End With

    Application.GoTo Sheets("System Selection").Range("A1"), True

    With Sheets("System Selection")
        If .Range("B4").Value = "" Then
            .Range("B4").Select
            customerInput = InputBox("Please fill in the Customer Name.", "Customer Name")
            ActiveCell.FormulaR1C1 = customerInput
        End If

        fileSaveName = "IBU Inventory BOM" & ".xlsm"
        rigSaveName = CleanFileName(.Range("Z4").Value) & " - "
        customerSaveName = CleanFileName(.Range("B4").Value) & " - "
    End With

 'Set the Root Path

    fileRootPath = ThisWorkbook.Path & "\"

 'Set the sub paths

    fileSavePath = fileRootPath & customerSaveName & rigSaveName

    ActiveWorkbook.SaveAs Filename:=fileSavePath & fileSaveName

End Function

Function UnhideWorksheets()

    With ActiveWorkbook
    
    .Worksheets("Rig Survey Form").Visible = True
    .Worksheets("System Selection").Visible = True
    .Worksheets("General Use Items").Visible = True
    .Worksheets("Master Parts List").Visible = True
        
    End With

End Function

Function CleanFileName(sFileName As String, Optional ReplaceInvalidwith As String = "") As String
    'Removes invalid filename characters

    Const InvalidChars As String = "%~:\/?*<>|"""
    Dim ThisChar As Long
    CleanFileName = sFileName
    For ThisChar = 1 To Len(InvalidChars)
        CleanFileName = Replace(CleanFileName, Mid(InvalidChars, ThisChar, 1), ReplaceInvalidwith)
    Next
End Function

Function SaveWorkbook()
'Created for Pason by Dragon Wood (October 2015).
'Saves the workbook.

    ThisWorkbook.Save
    
End Function

Function ShowLegend()
'Created for Pason by Dragon Wood (October 2015).
'Displays the Legend Form.
    
    Dim LegendControl As Boolean
    
        LegendControl = ActiveWorkbook.Worksheets("Rig Survey Form").Range("AG12")
    
    If LegendControl = False Then
        frmLegend.Show vbModeless
    Else
        Exit Function
    End If
    
End Function

Function CloseLegend()

    Unload frmLegend
    
End Function

Function ResetLegend()

    Dim LegendControl As Boolean
    
        LegendControl = ActiveWorkbook.Worksheets("Rig Survey Form").Range("AG12")
    
    If LegendControl = True Then
        ActiveWorkbook.Worksheets("Rig Survey Form").Range("AG12").Value = False
    Else
        Exit Function
    End If

End Function

Function RemapEnterKey()
    'numeric keypad
    Call Application.OnKey("{ENTER}", "NewEnter")
    'regular Enter
    Call Application.OnKey("~", "NewEnter")
End Function

Function UnmapEnterKey()
    'numeric keypad
    Call Application.OnKey("{ENTER}")
    'regular Enter
    Call Application.OnKey("~")
End Function

Function NewEnter()
    If Not TypeOf Selection Is Range Then Exit Function
    'Application.SendKeys "{TAB}"
    Dim keys As New clsKeyboard
    keys.PressKeyVK keyTab
End Function

Function Test_clsKeyboard()
  Dim keys As New clsKeyboard
  keys.PressKeyVK keyTab
End Function

Public Sub SetSystemSelectionBoxes()
     'Sets the options for the drop boxes on the System Selection sheet.
    With Worksheets("System Selection")
        With .VoltageBox
            .AddItem "110V"
            .AddItem "220V"
        End With
        With .UnitsBox
            .AddItem "Imperial"
            .AddItem "Metric"
            .AddItem "Mixed"
        End With
        With .TorqueBox
            .AddItem "Electric"
            .AddItem "Hydraulic"
        End With
        With .ProbeBox
            .AddItem "Radar"
            .AddItem "Mud Probe"
            .AddItem "Both"
        End With
        For Each itm In Array(.FlowBox, .UJBBox)
            itm.AddItem "Yes"
            itm.AddItem "No"
        Next itm
        For Each itm In Array(.AutoDrillerBox, .CasingBox, .ChokeBox, .EDRBox, .ePVTBox, .ESRBox, .GABox, .HGasBox, .PRDBox, .PVTBox, .SideKickBox, .WorkstationsBox)
            itm.AddItem "No"
            itm.AddItem "Yes"
        Next itm
    End With

End Sub

Function NewProjectMessage()
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