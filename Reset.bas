Attribute VB_Name = "Module3"
Sub ResetButton()

    'Confirmation Box:
    Dim Affirmed As Integer
    Affirmed = MsgBox("Resetting the Headcount Sheet cannot be undone. Proceed?", vbYesNo, "Delete A Days Data")
    If Affirmed = vbYes Then
        Reset
    End If

End Sub

Sub Reset()
    
    'Resetting Color:
    WIPColor = Range("R1").Interior.Color 'Note that putting a format on R1 will ruin the function of this button.
    WIPColor = Range("R1").Interior.Color 'Note that putting a format on R1 will ruin the function of this button.
    Range("A1:L10").Interior.Color = WIPColor
    Range("A11:F11").Interior.Color = WIPColor
    Range("A13:F31").Interior.Color = WIPColor
    Range("E32:E33").Interior.Color = WIPColor
    Range("C33:F70").Interior.Color = WIPColor
    Range("D71:F77").Interior.Color = WIPColor
    Range("C78:F5000").Interior.Color = WIPColor
    Range("A33:B5000").Interior.Color = WIPColor
    Range("H11:J5000").Interior.Color = WIPColor
    Range("K1:K70").Interior.Color = WIPColor
    Range("L19:N5000").Interior.Color = WIPColor
    
    'Resetting Values:
    Range("Q2:Q18").Value = 0
    Range("S2:S18").Value = ""

    'Important Constants:
    PAM = "Pattern Analysis Generator"
    WSName = Range("U6").Value
    CurrentSearchRange = "A1:A5000" 'The area of the PAM worksheet to search for the Observation Name to confirm no redundancies. In the theoretical case that the number is above 5000, the logic that checks for redundant names WILL stop working.

    'Warning message, related to CSR constant:
    Dim CNT As Integer
    CNT = Worksheets(PAM).Range("B1").Value
    If CNT > 4900 Then
        WarningBox = MsgBox("WARNING: If you exceed 5000 patterner sheets, this program WILL NOT work without modification.", vbOKOnly)
    End If

    'Resetting Patterner:

    If Worksheets(PAM).Range(CurrentSearchRange).Find(WSName) Is Nothing Then
    
        InitiatePatterner 'This module is currently only called in this action button.
    
    Else
        WarningBox = MsgBox("Observation Name already in use, refreshing previous patterner.", vbOKOnly)
        Worksheets(WSName).Range("A1:AA5000").Clear
    End If
        

End Sub
