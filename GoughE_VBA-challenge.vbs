Sub TickerSummary():
    
    ' Loop through all sheets
    Dim ws As Worksheet
    For Each ws In Worksheets

        ' Find the Last Row
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Put in some headers for summary tables
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
    
        ' Define everything
        Dim ticker As String
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Stock_Volume As Double
        Dim Open_Value As Double
        Dim Close_Value As Double
        Dim Year_Open As String
        Year_Open = ws.Name & "0102"
        ' MsgBox ("the year open is " & Year_Open)
        Dim Year_Close As String
        Year_Close = ws.Name & "1231"
        ' MsgBox ("the year close is " & Year_Close)
    
        ' Set an initial variable for holding the total stock volume
        Stock_Volume = 0
    
        ' Keep Track of the location for each ticker in the summary table
        Dim Summary_Table_Row As Double
        Summary_Table_Row = 2
    
        ' Loop through all tickers
        For I = 2 To LastRow
            
            ' Check if we are still within the same ticker, if it is not...
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
            
                ' Set the Ticker name
                ticker = ws.Cells(I, 1).Value
                
                ' Nested if for Close Value
                If ws.Cells(I, 2).Value = Year_Close Then
                
                    Close_Value = ws.Cells(I, 6).Value
                    
                End If
                
                ' Add to the Stock Volume
                Stock_Volume = Stock_Volume + ws.Cells(I, 7).Value
                
                ' Set the Yearly Change
                Yearly_Change = Close_Value - Open_Value
                
                ' Set the Percent Change
                Percent_Change = (Close_Value / Open_Value) - 1
                
                ' Print the Ticker in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = ticker
                
                ' Print the Yearly Change in the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = Round(Yearly_Change, 2)
                
                ' Print the Percent Change in the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = FormatPercent(Percent_Change, 2)
                
                ' Print the Stock Volume to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
                
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Reset the Stock Volume
                Stock_Volume = 0
                
            ' If the cell immediately following a row is the same ticker...
            Else
            
                ' Nested If for Open_value
                If ws.Cells(I, 2).Value = Year_Open Then
                
                    Open_Value = ws.Cells(I, 3).Value
                    
                End If
                
                ' Add to the Stock Volume
                Stock_Volume = Stock_Volume + ws.Cells(I, 7).Value
                
            End If
            
        Next I
        
        ' Find the last row of the summary table
        Dim SumLastRow As Long
        SumLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        ' Define variables for second summary table
        Dim TopPercent As Double
        Dim BottomPercent As Double
        Dim TopVolume As Double
    
        ' Calculate Values
        TopPercent = WorksheetFunction.Max(ws.Range("K2", "K" & SumLastRow))
        ' MsgBox ("The max is " & TopPercent)
        BottomPercent = WorksheetFunction.Min(ws.Range("K2", "K" & SumLastRow))
        ' MsgBox ("The min is " & BottomPercent)
        TopVolume = WorksheetFunction.Max(ws.Range("L2", "L" & SumLastRow))
        ' MsgBox ("The max is " & TopVolume)
    
        ' Print Values
        ws.Range("Q2").Value = FormatPercent(TopPercent, 2)
        ws.Range("Q3").Value = FormatPercent(BottomPercent, 2)
        ws.Range("Q4").Value = TopVolume
    
        ' Create Loop for Second Summary Table
            For I = 2 To SumLastRow
            
                ' Conditional Formatting
                If Cells(I, 10).Value >= 0 Then
            
                    ws.Cells(I, 10).Interior.ColorIndex = 4
                
                Else
            
                    ws.Cells(I, 10).Interior.ColorIndex = 3
                
                End If
            
                ' Print Values
                If ws.Cells(I, 11) = TopPercent Then
                    ticker = ws.Cells(I, 9).Value
                    ws.Range("P2").Value = ticker
                End If
                 
                If ws.Cells(I, 11) = BottomPercent Then
                    ticker = ws.Cells(I, 9).Value
                    ws.Range("P3").Value = ticker
                End If
                
                If Cells(I, 12) = TopVolume Then
                    ticker = ws.Cells(I, 9).Value
                    ws.Range("P4").Value = ticker
                
                End If
            
            Next I

        ' AutoFit the columns
        ws.Cells.EntireColumn.AutoFit

    Next ws

End Sub

