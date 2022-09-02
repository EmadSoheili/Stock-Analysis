Sub MacroCheck()

    Dim testMessage As String
    testMessage = "Hello World!"
    MsgBox (testMessage)
    
End Sub

Sub DQAnalysis()

    Worksheets("DQ Analysis").Activate
    
    Range("A1").Value = "DAQO (Ticker: DQ)"
    
    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
        
    '========================================
        
    Worksheets("2018").Activate
    
    RowStart = 2
    RowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    'RowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
    
    TotalVolume = 0
    Dim StartingPrice As Double
    Dim EndingPrice As Double
    
    For i = RowStart To RowEnd
        
        If Cells(i, 1).Value = "DQ" Then
            TotalVolume = TotalVolume + Cells(i, 8).Value
        End If
        
        If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then
            StartingPrice = Cells(i, 6).Value
        End If
        
        If Cells(i, 1).Value <> "DQ" And Cells(i - 1, 1).Value = "DQ" Then
            EndingPrice = Cells(i - 1, 6).Value
        End If
        
    Next i

    Worksheets("DQ Analysis").Activate

    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = TotalVolume
    Cells(4, 3).Value = EndingPrice / StartingPrice - 1

End Sub


Sub AllStocksAnalysis()

    'Step 1: Format the Output sheet ========================================

        Worksheets("All Stocks Analysis").Activate
        
        Dim StartTime As Single
        Dim EndTime As Single
                
        YearValue = InputBox("What year would you like to run the analysis on?")
        
        StartTime = Timer
        
        Range("A1").Value = "All Stocks (" + YearValue + ")"
        
        'Create a header row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"
        
        
    'Step 2: Create an array of all Tickers =================================
        
        Dim tickers(11) As String

        tickers(0) = "AY"
        tickers(1) = "CSIQ"
        tickers(2) = "DQ"
        tickers(3) = "ENPH"
        tickers(4) = "FSLR"
        tickers(5) = "HASI"
        tickers(6) = "JKS"
        tickers(7) = "RUN"
        tickers(8) = "SEDG"
        tickers(9) = "SPWR"
        tickers(10) = "TERP"
        tickers(11) = "VSLR"
        
    'Step 3: Prepare for Analysis ===========================================
        
        ' Activating Worksheet
        Worksheets(YearValue).Activate
        
        ' Creating Variables
        RowStart = 2
        RowEnd = Cells(Rows.Count, "A").End(xlUp).Row
            'RowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
        Dim StartingPrice As Double
        Dim EndingPrice As Double
        
    'Step 4: Loop through the Tickers =======================================
        
        For i = 0 To 11
            
            ticker = tickers(i)
            TotalVolume = 0
            
            Worksheets(YearValue).Activate
            
    'Step 5: Loop through the Rows ==========================================
        
            For j = RowStart To RowEnd
        
                If Cells(j, 1).Value = ticker Then
                    TotalVolume = TotalVolume + Cells(j, 8).Value
                End If
                
                If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                    StartingPrice = Cells(j, 6).Value
                End If
                
                If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    EndingPrice = Cells(j, 6).Value
                End If
            
            Next j
                   
    'Step 6: Output the findings ============================================

            Worksheets("All Stocks Analysis").Activate

            Cells(i + 4, 1).Value = ticker
            Cells(i + 4, 2).Value = TotalVolume
            Cells(i + 4, 3).Value = EndingPrice / StartingPrice - 1
            
        Next i
        
        EndTime = Timer
        MsgBox "This code ran in " & (EndTime - StartTime) & " seconds for the year " & YearValue
            
End Sub

Sub formatAllStocksAnalysisTable()

    Worksheets("All Stocks Analysis").Activate
    
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Range("B4:B15").NumberFormat = "#,##0"
    Range("c4:c15").NumberFormat = "0.00%"
    Columns("A:C").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd
    
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
        
        ElseIf Cells(i, 3) < 0 Then
            Cells(i, 3).Interior.Color = vbRed
            
        Else
            Cells(i, 3).Interior.Color = xlNone
        End If
        
    Next i
    
End Sub

Sub ClearWorksheet()

    Cells.Clear

End Sub
