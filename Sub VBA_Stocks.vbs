Sub VBA_Stocks()
Dim Rnd As String
Dim NR, NC, row_idx, NWS As Long
Dim OP_PR, CL_PR, val As Double
'  number of worksheets in the workbook
NWS = ActiveWorkbook.Worksheets.Count
For k = 1 To NWS
    'Activating the worksheet k
    Worksheets(k).Activate
    'telling to the user the name of the worksheet I am currently working on
    MsgBox (ActiveWorkbook.Worksheets(k).Name)
    'evaluating the total number of rows containing data
    NR = ActiveSheet.UsedRange.Rows.Count
    'MsgBox (NR)
    'evaluating the total number of columns containing data
    NC = ActiveSheet.UsedRange.Columns.Count ' - How can I fix this if I have an empty row between two set of rows?
    'MsgBox (NC)
    If NC > 7 Then
        NC = 7
    End If
    'external counter and variables
    Rnd = Cells(2, 3).Value
    row_idx = 2
    'loops through all the rows
    For i = 2 To NR + 1
        ticker = Cells(i, 1).Value
        If Rnd <> ticker Then
        'writing the ticker
            Cells(row_idx, NC + 2).Value = ticker
        'writing the total volume for the previous ticker
            If row_idx <> 2 Then
                'writing the $ yearly difference for the previous ticker
                Cells(row_idx - 1, NC + 3).Value = CL_PR - OP_PR
                'writing the % difference for the previous ticker
                With Cells(row_idx - 1, NC + 4)
                    If OP_PR = CL_PR And OP_PR = 0 Then
                        MsgBox ("WARNING: " + ticker + " was zero at the beginning and end of the year")
                        .Value = 0
                        .Interior.Color = vbRed
                    ElseIf OP_PR <> CL_PR And OP_PR = 0 Then
                        MsgBox ("WARNING: " + ticker + " was zero at the beginning of the year but closed at " + CStr(CL_PR))
                        .Value = "N/A"
                        If CL_PR > 0 Then
                            .Interior.Color = vbGreen
                        ElseIf CL_PR < 0 Then
                        .Interior.Color = vbRed
                        End If
                    Else
                        .Value = (CL_PR - OP_PR) / OP_PR
                        .NumberFormat = "0.00%"
                        If (CL_PR - OP_PR) / OP_PR < 0 Then
                            .Interior.Color = vbRed
                        Else
                            .Interior.Color = vbGreen
                        End If
                    End If
                End With
                'writing the total volume for the previous ticker
                Cells(row_idx - 1, NC + 5).Value = volume
            End If
            'saving the opening price
            OP_PR = Cells(i, 3).Value
            'updating external counter and controls
            Rnd = Cells(row_idx, 1 + NC + 1).Value
            row_idx = row_idx + 1
            'saving the volume for the first day
            volume = Cells(i, NC)
        Else
            volume = volume + Cells(i, NC)
                'overwriting the final price for the year for every ietartion
            CL_PR = Cells(i, 6).Value
        End If
    Next
    
    'setting up the titles of the columns to populate
    Cells(1, NC + 2).Value = "Ticker"
    Cells(1, NC + 3).Value = "Yearly price change"
    Cells(1, NC + 4).Value = "Yearly % change"
    Cells(1, NC + 5).Value = "Total Volume"
    
    ' evaluating the maximum and minimum values with respect to the ticker
    'MsgBox (row_idx)
    Cells(1, NC + 8).Value = "Value"
    Cells(1, NC + 9).Value = "Ticker"
    Greatest_array = Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume")
    For i = 0 To 2
        Cells(i + 2, NC + 7).Value = Greatest_array(i)
        If Greatest_array(i) = "Greatest % Increase" Then
            val = Application.WorksheetFunction.Max(Range("K:K"))
        ElseIf Greatest_array(i) = "Greatest % Decrease" Then
            val = Application.WorksheetFunction.Min(Range("K:K"))
        Else '"Greatest Total Volume"
            val = Application.WorksheetFunction.Max(Range("L:L"))
        End If
        Cells(i + 2, NC + 8).Value = val
        With Cells(i + 2, NC + 8)
        .Value = val
        For j = 2 To row_idx + 1
            If Cells(j, 11).Value = val And (Greatest_array(i) = "Greatest % Increase" Or Greatest_array(i) = "Greatest % Decrease") Then
                Cells(i + 2, NC + 9).Value = Cells(j, 9).Value
                .NumberFormat = "0.00%"
            ElseIf Cells(j, 12).Value = val And Greatest_array(i) = "Greatest Total Volume" Then
                Cells(i + 2, NC + 9).Value = Cells(j, 9).Value
                .NumberFormat = "0.000E+0"
            End If
        Next j
        End With
    Next i
Next k
        
End Sub

