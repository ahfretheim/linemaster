Attribute VB_Name = "Module2"
'VERSION NOTE: As it stands right now, this code will allow you to correct a mistake, but not undo a correction.


Sub Macro1()
Attribute Macro1.VB_Description = "Automates a Hotkey"
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' Macro1 Macro
' Automates a Hotkey
'
' Keyboard Shortcut: Ctrl+q
'
    'Patterner Functionality:
    
    Pattern 1
    
    '.Interior.Color captures the exact color property of any cell in an excel spreadsheet. (Other Excel objects also have Interior properties, and it can be used there as well.)
    
    ActiveCell.Interior.Color = Range("M2").Interior.Color
    
    'VBA does not have a += operator, so we will need to use a storage variable as an intermediate holder of the value
    
    Dim cellRecorded As String
    cellRecorded = CStr(ActiveCell.Row()) & ", " & CStr(ActiveCell.Column) & ";"
    
    If Range("S2").Find(cellRecorded) Is Nothing Then
        Dim Storage As String
        Dim rng As String
        rng = "S3:S18"
        Storage = Range("S2").Value
        Range("S2").Value = Storage & " " & cellRecorded
        Dim CurrentValue As Integer
        CurrentValue = Range("Q2").Value
        Range("Q2").Value = CurrentValue + 1
        ReplaceWatch cellRecorded, rng
        
        
        
    End If


End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = "w\n14"
'
' Macro2 Macro
'
' Keyboard Shortcut: Ctrl+w
'

    'Patterner Functionality:
    Pattern 2

    '.Interior.Color captures the exact color property of any cell in an excel spreadsheet. (Other Excel objects also have Interior properties, and it can be used there as well.)
    
    ActiveCell.Interior.Color = Range("M3").Interior.Color
    
    'VBA does not have a += operator, so we will need to use a storage variable as an intermediate holder of the value
    
    Dim cellRecorded As String
    cellRecorded = CStr(ActiveCell.Row()) & ", " & CStr(ActiveCell.Column) & ";"
    
    
    If Range("S3").Find(cellRecorded) Is Nothing Then
        Dim Storage As String
        Storage = Range("S3").Value
        Range("S3").Value = Storage & " " & cellRecorded
        Dim CurrentValue As Integer
        CurrentValue = Range("Q3").Value
        Range("Q3").Value = CurrentValue + 1
        ReplaceWatch cellRecorded, "S2"
        ReplaceWatch cellRecorded, "S4:S18"
    End If

End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = "e\n14"
'
' Macro3 Macro
'
' Keyboard Shortcut: Ctrl+e
'
    'Patterner functionality:
    
    Pattern 3

    '.Interior.Color captures the exact color property of any cell in an excel spreadsheet. (Other Excel objects also have Interior properties, and it can be used there as well.)
    
    ActiveCell.Interior.Color = Range("M4").Interior.Color
    
    'VBA does not have a += operator, so we will need to use a storage variable as an intermediate holder of the value
    
    Dim cellRecorded As String
    cellRecorded = CStr(ActiveCell.Row()) & ", " & CStr(ActiveCell.Column) & ";"
    
    If Range("S4").Find(cellRecorded) Is Nothing Then
        Dim Storage As String
        Storage = Range("S4").Value
        Range("S4").Value = Storage & " " & cellRecorded
        Dim CurrentValue As Integer
        CurrentValue = Range("Q4").Value
        Range("Q4").Value = CurrentValue + 1
        ReplaceWatch cellRecorded, "S2:S3"
        ReplaceWatch cellRecorded, "S5:S18"
    End If

End Sub
Sub Macro0()
Attribute Macro0.VB_ProcData.VB_Invoke_Func = "p\n14"
'NOTE: Whenever you add additional values above the 0 key, you'll need to manually adjust the cell references in this macro. They DO NOT adjust automatically.


'
' Macro0 Macro
'
' Keyboard Shortcut: Ctrl+p
'
    'Patterner functionality:
    Pattern 0


    '.Interior.Color captures the exact color property of any cell in an excel spreadsheet. (Other Excel objects also have Interior properties, and it can be used there as well.)
    
    ActiveCell.Interior.Color = Range("M8").Interior.Color
    
    'VBA does not have a += operator, so we will need to use a storage variable as an intermediate holder of the value
    

    
    'Error corecting code for situations where the hotkeys gets pushed multiple times in the same cell:
    
    Dim cellRecorded As String
    cellRecorded = CStr(ActiveCell.Row()) & ", " & CStr(ActiveCell.Column) & ";"
    
    If Range("S8").Find(cellRecorded) Is Nothing Then
        Dim Storage As String
        Storage = Range("S8").Value
        Range("S8").Value = Storage & " " & cellRecorded
        Dim CurrentValue As Integer
        CurrentValue = Range("Q8").Value
        Range("Q8").Value = CurrentValue + 1
        ReplaceWatch cellRecorded, "S2:S7"
        ReplaceWatch cellRecorded, "S9:S18"
    End If

End Sub
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = "r\n14"
'
' Macro4 Macro
'
' Keyboard Shortcut: Ctrl+r
'

    'Patterner Functionality:
    
    Pattern 4

    ActiveCell.Interior.Color = Range("M5").Interior.Color
    
    'VBA does not have a += operator, so we will need to use a storage variable as an intermediate holder of the value
    
    Dim cellRecorded As String
    cellRecorded = CStr(ActiveCell.Row()) & ", " & CStr(ActiveCell.Column) & ";"
    
    If Range("S5").Find(cellRecorded) Is Nothing Then
        Dim Storage As String
        Storage = Range("S5").Value
        Range("S5").Value = Storage & " " & cellRecorded
        Dim CurrentValue As Integer
        CurrentValue = Range("Q5").Value
        Range("Q5").Value = CurrentValue + 1
        ReplaceWatch cellRecorded, "S2:S4"
        ReplaceWatch cellRecorded, "S6:S18"
        
    End If

End Sub
Sub Macro6()
Attribute Macro6.VB_ProcData.VB_Invoke_Func = "y\n14"
'
' Macro6 Macro
'
' Keyboard Shortcut: Ctrl+y
'

    'Patterner Functionality:
    
    Pattern 6

    ActiveCell.Interior.Color = Range("M6").Interior.Color
    
    'VBA does not have a += operator, so we will need to use a storage variable as an intermediate holder of the value
    
    Dim cellRecorded As String
    cellRecorded = CStr(ActiveCell.Row()) & ", " & CStr(ActiveCell.Column) & ";"
    
    If Range("S6").Find(cellRecorded) Is Nothing Then
        Dim Storage As String
        Storage = Range("S6").Value
        Range("S6").Value = Storage & " " & cellRecorded
        Dim CurrentValue As Integer
        CurrentValue = Range("Q6").Value
        Range("Q6").Value = CurrentValue + 1
        ReplaceWatch cellRecorded, "S2:S5"
        ReplaceWatch cellRecorded, "S7:S18"
    End If
End Sub
Sub Macro7()
Attribute Macro7.VB_ProcData.VB_Invoke_Func = "u\n14"
'
' Macro7 Macro
'
' Keyboard Shortcut: Ctrl+u
'

    'Patterner functionality:
    
    Pattern 7

    ActiveCell.Interior.Color = Range("M7").Interior.Color
    
    'VBA does not have a += operator, so we will need to use a storage variable as an intermediate holder of the value
    
    Dim cellRecorded As String
    cellRecorded = CStr(ActiveCell.Row()) & ", " & CStr(ActiveCell.Column) & ";"
    
    If Range("S7").Find(cellRecorded) Is Nothing Then
        Dim Storage As String
        Storage = Range("S7").Value
        Range("S7").Value = Storage & " " & cellRecorded
        Dim CurrentValue As Integer
        CurrentValue = Range("Q7").Value
        Range("Q7").Value = CurrentValue + 1
        ReplaceWatch cellRecorded, "S2:S6"
        ReplaceWatch cellRecorded, "S8:S18"
    End If
End Sub
Sub MacroA()
Attribute MacroA.VB_ProcData.VB_Invoke_Func = "a\n14"
'
' MacroA Macro
'
' Keyboard Shortcut: Ctrl+a
'

    'Patterner functionality - current strategy for partial assignments is to leave them as letters in the patterner sheet and render in to fractions on the analysis-end:
    
    Pattern "A"

    ActiveCell.Interior.Color = Range("M9").Interior.Color
    
    'VBA does not have a += operator, so we will need to use a storage variable as an intermediate holder of the value
    
    Dim cellRecorded As String
    cellRecorded = CStr(ActiveCell.Row()) & ", " & CStr(ActiveCell.Column) & ";"
    
    If Range("S9").Find(cellRecorded) Is Nothing Then
        Dim Storage As String
        Storage = Range("S9").Value
        Range("S9").Value = Storage & " " & cellRecorded
        Dim CurrentValue As Integer
        CurrentValue = Range("Q9").Value
        Range("Q9").Value = CurrentValue + 1
        ReplaceWatch cellRecorded, "S2:S8"
        ReplaceWatch cellRecorded, "S10:S18"
    End If
End Sub
Sub MacroS()
Attribute MacroS.VB_ProcData.VB_Invoke_Func = "s\n14"
'
' MacroS Macro
'
' Keyboard Shortcut: Ctrl+s
'
    'Patterner functionality - current strategy for partial assignments is to leave them as letters in the patterner sheet and render in to fractions on the analysis-end:
    Pattern "S"
 
    ActiveCell.Interior.Color = Range("M10").Interior.Color
    
    'VBA does not have a += operator, so we will need to use a storage variable as an intermediate holder of the value
    
    Dim cellRecorded As String
    cellRecorded = CStr(ActiveCell.Row()) & ", " & CStr(ActiveCell.Column) & ";"
    
    If Range("S10").Find(cellRecorded) Is Nothing Then
        Dim Storage As String
        Storage = Range("S10").Value
        Range("S10").Value = Storage & " " & cellRecorded
        Dim CurrentValue As Integer
        CurrentValue = Range("Q10").Value
        Range("Q10").Value = CurrentValue + 1 'This is changed to 2 to correct for the action of ReplaceWatch below
        ReplaceWatch cellRecorded, "S2:S9"
        ReplaceWatch cellRecorded, "S11:S18"
    End If
End Sub
Sub MacroD()
Attribute MacroD.VB_ProcData.VB_Invoke_Func = "d\n14"
'
' MacroD Macro
'
' Keyboard Shortcut: Ctrl+d
'

    'Patterner functionality - current strategy for partial assignments is to leave them as letters in the patterner sheet and render in to fractions on the analysis-end:
    Pattern "D"

    ActiveCell.Interior.Color = Range("M11").Interior.Color
    
    'VBA does not have a += operator, so we will need to use a storage variable as an intermediate holder of the value
    
    Dim cellRecorded As String
    cellRecorded = CStr(ActiveCell.Row()) & ", " & CStr(ActiveCell.Column) & ";"
    
    If Range("S11").Find(cellRecorded) Is Nothing Then
        Dim Storage As String
        Storage = Range("S11").Value
        Range("S11").Value = Storage & " " & cellRecorded
        Dim CurrentValue As Integer
        CurrentValue = Range("Q11").Value
        Range("Q11").Value = CurrentValue + 1
        ReplaceWatch cellRecorded, "S2:S10"
        ReplaceWatch cellRecorded, "S12:S18"
    End If
End Sub
Sub MacroF()
Attribute MacroF.VB_ProcData.VB_Invoke_Func = "f\n14"
'
' MacroF Macro
'
' Keyboard Shortcut: Ctrl+f
'

    'Patterner functionality - current strategy for partial assignments is to leave them as letters in the patterner sheet and render in to fractions on the analysis-end:
 
    Pattern "F"

    ActiveCell.Interior.Color = Range("M12").Interior.Color
    
    'VBA does not have a += operator, so we will need to use a storage variable as an intermediate holder of the value
    
    Dim cellRecorded As String
    cellRecorded = CStr(ActiveCell.Row()) & ", " & CStr(ActiveCell.Column) & ";"
    
    If Range("S12").Find(cellRecorded) Is Nothing Then
        Dim Storage As String
        Storage = Range("S12").Value
        Range("S12").Value = Storage & " " & cellRecorded
        Dim CurrentValue As Integer
        CurrentValue = Range("Q12").Value
        Range("Q12").Value = CurrentValue + 1
        ReplaceWatch cellRecorded, "S2:S11"
        ReplaceWatch cellRecorded, "S13:S18"
    End If

End Sub
Sub MacroG()
Attribute MacroG.VB_ProcData.VB_Invoke_Func = "g\n14"
'
' MacroG Macro
'
' Keyboard Shortcut: Ctrl+g
'

    'Patterner functionality - current strategy for partial assignments is to leave them as letters in the patterner sheet and render in to fractions on the analysis-end:
    Pattern "G"

    ActiveCell.Interior.Color = Range("M13").Interior.Color
    
    'VBA does not have a += operator, so we will need to use a storage variable as an intermediate holder of the value
    
    Dim cellRecorded As String
    cellRecorded = CStr(ActiveCell.Row()) & ", " & CStr(ActiveCell.Column) & ";"
    
    If Range("S13").Find(cellRecorded) Is Nothing Then
        Dim Storage As String
        Storage = Range("S13").Value
        Range("S13").Value = Storage & " " & cellRecorded
        Dim CurrentValue As Integer
        CurrentValue = Range("Q13").Value
        Range("Q13").Value = CurrentValue + 1
        ReplaceWatch cellRecorded, "S2:S12"
        ReplaceWatch cellRecorded, "S14:S18"
    End If
End Sub
Sub MacroH()
Attribute MacroH.VB_ProcData.VB_Invoke_Func = "h\n14"
'
' MacroH Macro
'
' Keyboard Shortcut: Ctrl+h
'

    'Patterner functionality - current strategy for partial assignments is to leave them as letters in the patterner sheet and render in to fractions on the analysis-end:
    Pattern "H"

    ActiveCell.Interior.Color = Range("M14").Interior.Color
    
    'VBA does not have a += operator, so we will need to use a storage variable as an intermediate holder of the value
    
    Dim cellRecorded As String
    cellRecorded = CStr(ActiveCell.Row()) & ", " & CStr(ActiveCell.Column) & ";"
    
    If Range("S14").Find(cellRecorded) Is Nothing Then
        Dim Storage As String
        Storage = Range("S14").Value
        Range("S14").Value = Storage & " " & cellRecorded
        Dim CurrentValue As Integer
        CurrentValue = Range("Q14").Value
        Range("Q14").Value = CurrentValue + 1
        ReplaceWatch cellRecorded, "S2:S13"
        ReplaceWatch cellRecorded, "S15:S18"
    End If
End Sub
Sub MacroJ()
Attribute MacroJ.VB_ProcData.VB_Invoke_Func = "j\n14"
'
' MacroJ Macro
'
' Keyboard Shortcut: Ctrl+j
'

    'Patterner functionality - current strategy for partial assignments is to leave them as letters in the patterner sheet and render in to fractions on the analysis-end:
    Pattern "J"

    ActiveCell.Interior.Color = Range("M15").Interior.Color
    
    'VBA does not have a += operator, so we will need to use a storage variable as an intermediate holder of the value
    
    Dim cellRecorded As String
    cellRecorded = CStr(ActiveCell.Row()) & ", " & CStr(ActiveCell.Column) & ";"
    
    If Range("S15").Find(cellRecorded) Is Nothing Then
        Dim Storage As String
        Storage = Range("S15").Value
        Range("S15").Value = Storage & " " & cellRecorded
        Dim CurrentValue As Integer
        CurrentValue = Range("Q15").Value
        Range("Q15").Value = CurrentValue + 1
        ReplaceWatch cellRecorded, "S2:S14"
        ReplaceWatch cellRecorded, "S16:S18"
    End If
End Sub
Sub MacroK()
Attribute MacroK.VB_ProcData.VB_Invoke_Func = "k\n14"
'
' MacroK Macro
'
' Keyboard Shortcut: Ctrl+k
'

    'Patterner functionality - current strategy for partial assignments is to leave them as letters in the patterner sheet and render in to fractions on the analysis-end:
    Pattern "K"

    ActiveCell.Interior.Color = Range("M16").Interior.Color
    
    'VBA does not have a += operator, so we will need to use a storage variable as an intermediate holder of the value
    
    Dim cellRecorded As String
    cellRecorded = CStr(ActiveCell.Row()) & ", " & CStr(ActiveCell.Column) & ";"
    
    If Range("S16").Find(cellRecorded) Is Nothing Then
        Dim Storage As String
        Storage = Range("S16").Value
        Range("S16").Value = Storage & " " & cellRecorded
        Dim CurrentValue As Integer
        CurrentValue = Range("Q16").Value
        Range("Q16").Value = CurrentValue + 1
        ReplaceWatch cellRecorded, "S2:S15"
        ReplaceWatch cellRecorded, "S17:S18"
    End If
End Sub
Sub MacroL()
Attribute MacroL.VB_ProcData.VB_Invoke_Func = "l\n14"
'
' MacroL Macro
'
' Keyboard Shortcut: Ctrl+l
'

    'Patterner functionality - current strategy for partial assignments is to leave them as letters in the patterner sheet and render in to fractions on the analysis-end:
    Pattern "L"

    ActiveCell.Interior.Color = Range("M17").Interior.Color
    
    'VBA does not have a += operator, so we will need to use a storage variable as an intermediate holder of the value
    
    Dim cellRecorded As String
    cellRecorded = CStr(ActiveCell.Row()) & ", " & CStr(ActiveCell.Column) & ";"
    
    If Range("S17").Find(cellRecorded) Is Nothing Then
        Dim Storage As String
        Storage = Range("S17").Value
        Range("S17").Value = Storage & " " & cellRecorded
        Dim CurrentValue As Integer
        CurrentValue = Range("Q17").Value
        Range("Q17").Value = CurrentValue + 1
        ReplaceWatch cellRecorded, "S2:S16"
        ReplaceWatch cellRecorded, "S18"
    End If
End Sub
Sub MacroZ()
Attribute MacroZ.VB_ProcData.VB_Invoke_Func = "z\n14"
'
' MacroZ Macro
'
' Keyboard Shortcut: Ctrl+z
'

    'Patterner functionality - current strategy for partial assignments is to leave them as letters in the patterner sheet and render in to fractions on the analysis-end:
    Pattern "Z"

    ActiveCell.Interior.Color = Range("M18").Interior.Color
    
    'VBA does not have a += operator, so we will need to use a storage variable as an intermediate holder of the value
    
    Dim cellRecorded As String
    cellRecorded = CStr(ActiveCell.Row()) & ", " & CStr(ActiveCell.Column) & ";"
    
    If Range("S18").Find(cellRecorded) Is Nothing Then
        Dim Storage As String
        Storage = Range("S18").Value
        Range("S18").Value = Storage & " " & cellRecorded
        Dim CurrentValue As Integer
        CurrentValue = Range("Q18").Value
        Range("Q18").Value = CurrentValue + 1
        ReplaceWatch cellRecorded, "S2:S17"
    End If
End Sub

Sub ReplaceWatch(cellRecorded As String, rng As String)
    'This Submodule corrects any values changed by the user.
    
    If Not Range(rng).Find(cellRecorded) Is Nothing Then
        Dim CurrentValue As Integer
        Dim ToReduce As Integer
        ToReduce = Range(rng).Find(cellRecorded).Row()
        CurrentValue = ActiveSheet.Cells(ToReduce, 17).Value
        ActiveSheet.Cells(ToReduce, 17).Value = CurrentValue - 1
    End If
    
End Sub

'This Submodule adds a new Patterner worksheet when called. Currently only called by the reset button, as original vision for this function seems to have been deprecated with the "Contains" functionality from earlier versions of VBA.
Sub InitiatePatterner()
    
    'The name of the Pattern Analysis Worksheet:
    PAM = "Pattern Analysis Generator"
    
    
    'Create and name the Patterner sheet:
    WSName = Range("U6").Value
    Worksheets.Add
    ActiveSheet.Name = WSName
    'ActiveSheet.Protect Password = "tanks"
    Worksheets("Template").Activate
    ActiveSheet.Range("R19").Value = 0
    
    'Update the worksheets list in the Pattern Recognition Operating Panel:
    Dim CurrentNumber As Integer
    CurrentNumber = Worksheets(PAM).Range("B1").Value + 2
    Worksheets(PAM).Cells(CurrentNumber, 1).Value = WSName
    Worksheets(PAM).Range("B1").Value = CurrentNumber - 1
    
End Sub

Sub Pattern(val)

    Row = ActiveCell.Row
    Col = ActiveCell.Column
    PatternerName = Range("U6").Value

    'Initiates the sheet if it does not exist
    'If Not Contains(Sheets, PatternerName) Then
    '    InitiatePatterner
    'End If
    
    'Worksheets(PatternerName).Unprotect Password = "tanks"
    
    Worksheets(PatternerName).Cells(Row, Col).Value = val
    
End Sub
