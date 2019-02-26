Attribute VB_Name = "Module1"
Sub StockMarket()
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    '
    'GENERATION OF THE FIRST SUMMARY TABLE FOR MULTIPLE_YEAR_STOCK_DATA
    'labels summarized in the first table are "Ticker", "Yearly Change", "Percent Change", and "Total Stock Volume"
    '
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    Dim ws As Worksheet

    'loop through all worksheets
    For Each ws In Worksheets
    
        '---------------------------------------------
        'define original data ranges in each worksheet
        '---------------------------------------------
        
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        '---------------------------------------------
        'add headers (labels) to the 1st summary table
        '---------------------------------------------
        
        'store header names in an array called Headers
        Dim Headers() As String
        Headers = Split("Ticker,Yearly Change,Percent Change,Total Stock Volume", ",")
        
        'use x to iterate to add headers to the top row of the 1st summary table
        Dim x As Integer
        For x = 0 To 3
            ws.Cells(1, 9 + x).Value = Headers(x)
        Next x
        
        '----------------------------------------
        'variants setup for the 1st summary table
        '----------------------------------------
        
        'use SumTable1Row to locate the rows in the 1st Summary Table where sorted data of each ticker is stored
        'note that ticker's data start to be stored from Row 2
        Dim SumTable1Row As Integer
        SumTable1Row = 2
        
        'use PriceOpen to store each ticker's first open price of the year during iteration
        'note that ticker count is arranged chronologically, the value in Cell C2 is actually the open price for the first ticker in each worksheet
        Dim PriceOpen As Double
        PriceOpen = CDbl(ws.Cells(2, 3).Value)
        
        'use PriceClose to store each ticker's last closed price of the year during iteration
        Dim PriceClose As Double
              
        'use TotalStockVol to store the stock volume of each ticker during iteration
        Dim TotalStockVol As Double
        TotalStockVol = 0
                
        '----------------------------------------------------------------
        'Data extracting, analyzing, and writing to the 1st summary table
        '----------------------------------------------------------------
        
        'loop through stock data in each worksheet
        Dim i As Long
        For i = 2 To LastRow

            'add up ticker's stock volume during iteration
            'note that the stock volume of current iteration will be added to TotalStockVol no matter whether the next iteration has the same ticker or not
            TotalStockVol = TotalStockVol + ws.Cells(i, 7).Value

            'determine if the next iteration has the same ticker as the current one, and if not current iteration is the last data row for the ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'print the ticker in the 1st summary table
                ws.Cells(SumTable1Row, 9).Value = ws.Cells(i, 1).Value
                
                'print the ticker's total stock volume in the 1st summary table
                ws.Cells(SumTable1Row, 12).Value = TotalStockVol
                
                'store the ticker's last closed price of the year to PriceClose
                PriceClose = CDbl(ws.Cells(i, 6).Value)
                    
                'calculate and print the ticker's yearly change in the 1st summary table
                ws.Cells(SumTable1Row, 10).Value = PriceClose - PriceOpen
                'check if ticker's yearly change is less than 0
                If ws.Cells(SumTable1Row, 10).Value < 0 Then
                    'highlight the negative changes in red
                    ws.Cells(SumTable1Row, 10).Interior.ColorIndex = 3
                'check if ticker's yearly change is greater than 0
                ElseIf ws.Cells(SumTable1Row, 10).Value > 0 Then
                    'highlight the positive changes in green
                    ws.Cells(SumTable1Row, 10).Interior.ColorIndex = 4
                'should there be any yearly change ending up exactly with 0, it would not be highlighted
                End If

                'check if the ticker's open value of the year is zero, and if not ...
                If PriceOpen <> 0 Then
                    'calculate and print the ticker's percent change in the 1st summary table, set the value as "Percent"
                    ws.Cells(SumTable1Row, 11).Value = Format(ws.Cells(SumTable1Row, 10).Value / PriceOpen, "Percent")
                Else
                    'if the ticker's first open value of the year is 0, mark its percent change as not applicable (N/A)
                    ws.Cells(SumTable1Row, 11).Value = "N/A"
                End If
                                                                
                'if current iteration is not the bottom one of the original data ...
                'note that SumTable1Row will be referred to for the generation of 2nd summary table
                'a logically accurate SumTable1Row is appreciated, although otherwise it does not affect the output in 2nd summary table anyway
                If Not (i = LastRow) Then
                    'pass the open value from the next iteration to PriceOpen for the next stock ticker
                    PriceOpen = CDbl(ws.Cells(i + 1, 3).Value)
                    
                    'add 1 to SumTable1Row in the 1st summary table for data of the next stock ticker
                    SumTable1Row = SumTable1Row + 1
                
                    'reset the total of stock volume to 0 for the next stock ticker
                    TotalStockVol = 0
                End If
                                                
            End If
            
        Next i
            
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    '
    'GENERATION OF THE SECOND SUMMARY TABLE FOR MULTIPLE_YEAR_STOCK_DATA
    'labels summarized in the second table are "Ticker" and "Value" for columns and "Greatest % Increase", "Greatest %
    'Decrease", and "Greatest Total Volume" for rows
    'note that we are still in "for loop" for each ws
    '
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                        
        '-----------------------------------
        'add labels to the 2nd summary table
        '-----------------------------------
        
        'store label names in an array called Labels
        Dim Labels() As String
        Labels = Split("Greatest % Increase,Greatest % Decrease,Greatest Total Volume,Ticker,Value", ",")
        
        'add row labels for the 2nd summary table
        For x = 0 To 2
            ws.Cells(x + 2, 15).Value = Labels(x)
        Next x
        'add column labels for the 2nd summary table
        For x = 3 To 4
            ws.Cells(1, x + 13).Value = Labels(x)
        Next x
        
        '----------------------------------------
        'variants setup for the 2nd summary table
        '----------------------------------------
               
        'use MaxPctInc (greatest % increase), MaxPctDec (greatest % decrease), and MaxTotalStockVol (greatest total volume) to store maximum values during iteration
        Dim MaxPctInc, MaxPctDec, MaxTotalStockVol As Double
        
        'assign value from Cell K2 to MaxPctInc and MaxPctDec, set as their initial value in each worksheet
        MaxPctInc = CDbl(ws.Cells(2, 11).Value)
        MaxPctDec = CDbl(ws.Cells(2, 11).Value)
        
        'assign value from Cell L2 to MaxTotalStockVol, set as its initial value in each worksheet
        MaxTotalStockVol = CDbl(ws.Cells(2, 12).Value)
                        
        '----------------------------------------------------------------
        'Data extracting, analyzing, and writing to the 2nd summary table
        '----------------------------------------------------------------
                
        'loop through the 1st summary table
        For i = 3 To SumTable1Row
            
            'exclude ticker with 0 open value of the year
            If ws.Cells(i, 11).Value <> "N/A" Then
                
                'check if iterated percent change value is greater than MaxPctInc and if it is the case ...
                If ws.Cells(i, 11).Value > MaxPctInc Then
                    'replace MaxPctInc with current iterated value
                    MaxPctInc = ws.Cells(i, 11).Value
                    'write iterated ticker to Cell P2
                    ws.Range("P2").Value = ws.Cells(i, 9).Value
                    'write iterated percent change value to Cell Q2, set as "Percent"
                    ws.Range("Q2").Value = Format(ws.Cells(i, 11).Value, "Percent")
                    
                'if iterated percent change value is not greater than MaxPctInc, check if it is less than MaxPctDec (greater in absolute value) and if true ...
                ElseIf ws.Cells(i, 11).Value < MaxPctDec Then
                    'replace MaxPctDec with current iterated value
                    MaxPctDec = ws.Cells(i, 11).Value
                    'write iterated ticker to Cell P3
                    ws.Range("P3").Value = ws.Cells(i, 9).Value
                    'write iterated percent change value to Cell Q3, set the value as "Percent"
                    ws.Range("Q3").Value = Format(ws.Cells(i, 11).Value, "Percent")
                
                End If
            
            End If
            
            'check if iterated total stock volume is greater than MaxTotalStockVol and if it is the case ...
            If ws.Cells(i, 12).Value > MaxTotalStockVol Then
                'replace MaxTotalStockVol with current iterated value
                MaxTotalStockVol = ws.Cells(i, 12).Value
                'write iterated ticker to Cell P4
                ws.Range("P4").Value = ws.Cells(i, 9).Value
                'write iterated percent change value to Cell Q4
                ws.Range("Q4").Value = ws.Cells(i, 12).Value
            End If
        
        Next i
                                         
    Next ws
                
End Sub

            

