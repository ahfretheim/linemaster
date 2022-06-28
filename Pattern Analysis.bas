Attribute VB_Name = "Module4"
Sub PAMButton()

    'Initialization Variable Declarations:
    Dim PAMCount As Integer  'The Number of PAM Sheets created - used to uniquely name each sheet
    Dim IndexLimit As Integer 'Used to iterate over the patterners to create the Pattern Recognition
    Dim PAMName As String 'Name of the PAM Sheet
    
    'Initialization of the PAM Sheet:
    PAMCount = Range("C1").Value
    IndexLimit = Range("B1").Value + 1
    PAM = "Pattern Analysis Generator"
    CurrentSearchRange = "A1:A5000"
    WorkArea = "A1:M5000"

    'Adding the Pattern Analysis Worksheet
    Worksheets.Add
    PAMName = "Pattern Analysis " & PAMCount
    ActiveSheet.Name = PAMName
    
    Worksheets(PAM).Range("C1").Value = PAMCount + 1 'Ratchets the counter to prevent two PAM Sheets from having the same name
    
    'Row & Column Labels:
    ActiveSheet.Range(CurrentSearchRange).Value = Worksheets(PAM).Range(CurrentSearchRange).Value 'Row labels - Duplicates column A from PAM sheet
    ActiveSheet.Range("B1").Value = "Total Subassembly Workers" 'Column labels start here
    ActiveSheet.Range("C1").Value = "Total Workers at Assembly Stations"
    ActiveSheet.Range("D1").Value = "Most Workers at a Workstation"
    ActiveSheet.Range("E1").Value = "Idle Workstations"
    ActiveSheet.Range("F1").Value = "Most Stations Worked by a Single Worker"
    ActiveSheet.Range("G1").Value = "Number of Shared Workers"
    ActiveSheet.Range("H1").Value = "Average Stations Worked by a Shared Worker"
    ActiveSheet.Range("I1").Value = "Most Attended Station" 'the first 8 are summary statistics
    ActiveSheet.Range("J1").Value = "Most Changing Work Area" 'from this point forward, factors from deep analysis, that is, comparing multiple days. Each row is a comparison with all the rows before it.
    ActiveSheet.Range("K1").Value = "Unattended Stations Filled"
    'Currently just one deep factor, but the basic deep analysis structure, a 4-layered for loop, is complete.
    
    'Major Variable Declarations:
    Dim WorksheetName As String 'Container Variable for the Names of the Patterner Sheets
    Dim ToSumLeft As Range 'Half of Subassemblies
    Dim ToSumRight As Range 'Other Half of Subassemblies
    Dim Center As Range 'Assembly Stations
    Dim Workspace As Range 'The Entire Sheet
    Dim DeltName As String 'Name of the Delta Table worksheet
    
    'Labor Share Variable Declarations:
    
    'For every letter, every hotkey for every possible labor share in the template, the letter by itself is the total number of
    'stations the single worker is working, while the letter followed by 0 is whether the possible labor share is in use or not.
    'Therefore, it is always possible to calculate the total number of shared workers by summing all postscripted 0 variables.
    Dim A As Integer
    Dim A0 As Integer
    Dim S As Integer
    Dim S0 As Integer
    Dim D As Integer
    Dim D0 As Integer
    Dim F As Integer
    Dim F0 As Integer
    Dim G As Integer
    Dim G0 As Integer
    Dim H As Integer
    Dim H0 As Integer
    Dim J As Integer
    Dim J0 As Integer
    Dim K As Integer
    Dim K0 As Integer
    Dim L As Integer
    Dim L0 As Integer
    Dim Z As Integer
    Dim Z0 As Integer
    
    'Labor Share Helper Variables:
    Dim SharedLaborers As Integer
    Dim TotalSharedWorkStations As Integer
    Dim Capacity As Single         'Delta Table helper variable - needs to be a Double because Labor Shares result in harmonic fractions
    Dim Current As Single
    Dim Previous As Single
    Dim AbsoluteValue As Single
    
    'Locator Variable Declarations:
    Dim Supremum As Range
    Dim R As Integer
    Dim C As Integer
    Dim DiffSheet As String
    Dim DiffSheet2 As String
    Dim Filled As String 'A concatenated String used to hold & display all previously empty work areas from previous observations
    Dim ToFill As String 'Temporary container for concatenation to Filled. Overwritten with each use.
    Filled = "" 'DEBUG NOTE: Make sure to empty this variable again after each iterative use
    
    'Initializing the Delta Table for locating deep factors:
        Worksheets.Add
        DeltName = "Delta Table " & PAMCount
        ActiveSheet.Name = DeltName
        Worksheets(PAMName).Activate
    
    'Summary Statistics - calculated per observation:
    For i = 2 To IndexLimit
    
        'Selecting the patterner:
        WorksheetName = Cells(i, 1).Value
        Set Workspace = Worksheets(WorksheetName).Range("A1:N5000")
        Set Center = Worksheets(WorksheetName).Range("E1:I5000")
        
        'Finding Total Subassembly Workers for column B:
        Set ToSumLeft = Worksheets(WorksheetName).Range("A1:D5000")
        Set ToSumRight = Worksheets(WorksheetName).Range("J1:N5000")
        Cells(i, 2).Value = WorksheetFunction.Sum(ToSumLeft) + WorksheetFunction.Sum(ToSumRight)
    
        'Finding Maximum Workers for any workstation for column B:
        Cells(i, 4).Value = WorksheetFunction.Max(Workspace)
    
        'Finding Total Workers for the Main Line:
        Cells(i, 3).Value = WorksheetFunction.Sum(Center)
        
        'Finding Empty Workstations:
        Cells(i, 5).Value = WorksheetFunction.CountIf(Center, 0)
        
        'Labor Share-related Metrics:
        A = WorksheetFunction.CountIf(Workspace, "A")
        A0 = WorksheetFunction.Min(A, 1)
        S = WorksheetFunction.CountIf(Workspace, "S")
        S0 = WorksheetFunction.Min(S, 1)
        D = WorksheetFunction.CountIf(Workspace, "D")
        D0 = WorksheetFunction.Min(D, 1)
        F = WorksheetFunction.CountIf(Workspace, "F")
        F0 = WorksheetFunction.Min(F, 1)
        G = WorksheetFunction.CountIf(Workspace, "G")
        G0 = WorksheetFunction.Min(G, 1)
        H = WorksheetFunction.CountIf(Workspace, "H")
        H0 = WorksheetFunction.Min(H, 1)
        J = WorksheetFunction.CountIf(Workspace, "J")
        J0 = WorksheetFunction.Min(J, 1)
        K = WorksheetFunction.CountIf(Workspace, "K")
        K0 = WorksheetFunction.Min(K, 1)
        L = WorksheetFunction.CountIf(Workspace, "L")
        L0 = WorksheetFunction.Min(L, 1)
        Z = WorksheetFunction.CountIf(Workspace, "Z")
        Z0 = WorksheetFunction.Min(Z, 1)
        SharedLaborers = A0 + S0 + D0 + F0 + G0 + H0 + J0 + K0 + L0 + Z0
        TotalSharedWorkStations = A + S + D + F + G + H + J + K + L + Z
        Cells(i, 6).Value = WorksheetFunction.Max(A, S, D, F, G, H, J, K, L, Z)
        Cells(i, 7).Value = SharedLaborers
        
        If SharedLaborers > 0 Then
            Cells(i, 8).Value = TotalSharedWorkStations / SharedLaborers
        Else
            Cells(i, 8).Value = 0
        End If

        'Most Attended Station:
        
        Set Supremum = Center.Find(WorksheetFunction.Max(Center))
        R = Supremum.Row
        C = Supremum.Column
        Cells(i, 9).Value = Worksheets("Template").Cells(R, C).Value
        
        'Deep Factors - requires looking at the change between observations:
        
        'Note that while intermediate delta tables are created for every internal iterative step, only the final delta table remains for the user in the present version. This can be changed later if desired:
        
        If i = 2 Then
            'Blanks out all deep analysis Variables for first row
            ActiveSheet.Range("J2").Value = "N/A"
            ActiveSheet.Range("K2").Value = "N/A"
        Else
            'Initializing Delt Table & resetting Filled:
            
            Worksheets(DeltName).Range(WorkArea).Value = 0
            Filled = ""
        
            'Deep Analysis calculation For loop here:
            For J = 2 To i - 1    'Please note that iteration is used here because recursion in VBA is almost always a bad idea.
                'Level 2: sheet to sheet
                DiffSheet = Worksheets(PAM).Cells(J + 1, 1).Value
                DiffSheet2 = Worksheets(PAM).Cells(J, 1).Value
                'Calculation of Delt Table:
                For T = 1 To 5000
                    'Level 3: row index
                    For Y = 1 To 13
                        'Level 4: column index
                        'Calculating Delta Table:
                        Capacity = Worksheets(DeltName).Cells(T, Y).Value
                        Current = NumberRender(T, Y, DiffSheet)
                        Previous = NumberRender(T, Y, DiffSheet2)
                        AbsoluteValue = Abs(Current - Previous)
                        'Temporary debug code:
                        'If Current > 0 And Current < 1 Then
                        '    InfoBox = MsgBox("Current Value is " & Str(Current) & "Previous Value is " & Str(Previous) & " Delta is " + Str(AbsoluteValue), vbOKOnly)
                        'ElseIf Previous > 0 And Previous < 1 Then
                        '     InfoBox = MsgBox("Current Value is " & Str(Current) & "Previous Value is " & Str(Previous) & " Delta is " + Str(AbsoluteValue), vbOKOnly)
                        'End If
                            
                        Worksheets(DeltName).Cells(T, Y).Value = Capacity + AbsoluteValue
                        
                        'Temporary DebugCode:
                        'If InfoBox = 1 Then
                        '    InfoBox = MsgBox("Delta Table Value is " & Str(Worksheets(DeltName).Cells(T, Y).Value), vbOKOnly)
                        '    InfoBox = 0
                        'End If
                        
                        'Detecting Filled Previously Empty Work Areas:
                        If J = i - 1 Then
                            If Worksheets(DiffSheet2).Cells(T, Y).Value = 0 And Worksheets(DiffSheet).Cells(T, Y).Value <> 0 Then
                                ToFill = Worksheets("Template").Cells(T, Y).Value
                                Filled = Filled & ToFill & "; "
                            End If
                        End If
                    Next Y
                Next T
            Next J
            
            'Finding the Most Changed Work Area:
            Set Workspace = Worksheets(DeltName).Range("A1:N5000")
            Set Supremum = Workspace.Find(WorksheetFunction.Max(Workspace))
            R = Supremum.Row
            C = Supremum.Column
            Cells(i, 10).Value = Worksheets("Template").Cells(R, C).Value
            
            'Filling in the Previously Vacant Work Areas comparison column:
            If IsEmpty(Filled) Then
                Cells(i, 11).Value = "No Vacancies Filled"
            Else
                Cells(i, 11).Value = Filled
            End If
            
            
        End If

    Next i
    

   

    
End Sub

'This function is where partial labor assignments are handled for finding workstation deltas:
Function NumberRender(K, L, WS As String) As Single
    
    'D: Because this procedure is written as a function that returns a number, you CAN use it as a worksheet formula for troubleshooting purposes.
    Dim CheckForZeroes As Integer
    Dim ReturnValue As Single
    
    If IsNumeric(Worksheets(WS).Cells(K, L).Value) Then 'Numeric values always return whole numbers:
        NumberRender = Worksheets(WS).Cells(K, L).Value
    Else
        CheckForZeroes = WorksheetFunction.CountIf(Worksheets(WS).Range("A1:N5000"), Worksheets(WS).Cells(K, L).Value)
        'Temporary debugging code:
        'InfoBox = MsgBox("Value of CheckForZeroes is " & CheckForZeroes, vbOKOnly)
        If CheckForZeroes = 0 Then 'This code should never run
            NumberRender = 1
            PenaltyBox = MsgBox("ERROR: Number Render function CountIf not functioning correctly", vbOKOnly)
        ElseIf CheckForZeroes = 1 Then 'Temporary debugging code
            PenaltyBox = MsgBox("WARNING: CountIf only detecting one value for a labor share unit.", vbOKOnly)
            NumberRender = 1
        Else   'Labor shares return harmonic fractions:
            'Temporary debugging code:
            ReturnValue = 1 / CheckForZeroes
            'InfoBox = MsgBox("Harmonic Value returned as: " & ReturnValue, vbOKOnly)
            
            NumberRender = ReturnValue
        End If
    End If
End Function
