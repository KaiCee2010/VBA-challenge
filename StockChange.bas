Attribute VB_Name = "StockChange"
'Kaylon Young
'Stock Market Change Script
'Date: 2020-08-14
'This scripts notes the changes in stock market price over a given year
'each stock ticker


Sub StockChange()
        Dim ticker As String
        Dim openPrice, closePrice, yearlyChange, percentChange, totalVolume, minVal, maxVal, maxVolume As Double
        Dim rowNum As Integer
        Dim ws As Worksheet
        Dim myRange As Range
        Dim myRange2 As Range
        
        'Start worskheet loop
        For Each ws In ThisWorkbook.Worksheets
            'Activate worksheet
            ws.Activate
            Debug.Print "Started Worksheet " + ws.Name
            
            'Set variables
            ticker = Range("A2").Value
            rowNum = 2
            openPrice = Range("C2").Value
            closePrice = 0
            totalVolume = Range("G2").Value
            maxVal = 0
            
            'set worksheet values
            Range("I1").Value = "Ticker"
            Range("J1").Value = "Yearly Change"
            Range("K1").Value = "Percent Change"
            Range("L1").Value = "Stock Volume"
            Range("P1").Value = "Ticker"
            Range("Q1").Value = "Value"
            Range("O2").Value = "Greatest % Increase"
            Range("O3").Value = "Greatest % Decrease"
            Range("O4").Value = "Greatest Total Volume"
            
            'Find the last row of data in the worksheet
            lastrow = Cells(Rows.Count, 1).End(xlUp).Row
            Debug.Print "The last row in worksheet " + ws.Name + " is: " + Str(lastrow)
            
            'Begin to loop through the data
            For i = 2 To lastrow
                'determine if the ticker is the same
                If ticker <> Range("A" & i + 1) Then
                    'if same calculate yearly change
                    yearlyChange = closePrice - openPrice
                    
                    'calculate percent change, but only if values are not zero
                    If yearlyChange = 0 Or openPrice = 0 Then
                        percentChange = 0
                    Else
                        percentChange = yearlyChange / openPrice
                    End If
                    
                    'add ticket value to list
                    Range("I" & rowNum).Value = ticker
                    
                    'add yearly change value to the list
                    Range("J" & rowNum).Value = yearlyChange
                    
                    'change color of cell
                    If yearlyChange > 0 Then
                        Range("J" & rowNum).Interior.ColorIndex = 4
                    Else
                       Range("J" & rowNum).Interior.ColorIndex = 3
                    End If
                    
                    'add percent change to the list
                    Range("K" & rowNum).Value = FormatPercent(percentChange, 2)
                    
                    'add total volume to the list
                    Range("L" & rowNum).Value = totalVolume
                    
                    'reset my variables
                    ticker = Range("A" & i + 1).Value
                    openPrice = Range("C" & i + 1).Value
                    totalVolume = Range("G" & i + 1).Value
                    
                    'increment my row counts for my ticker list
                    rowNum = rowNum + 1
                    
                Else
                    'continually update the closing price
                    closePrice = Range("F" & i + 1).Value
                    
                    'continually add to the total volume
                    totalVolume = totalVolume + Range("G" & i + 1).Value
                    
                End If
                
            Next i
            
            'Starting to find min and max values
            'reset my row counter to the total number of rows in ticker list
            rowNum = rowNum - 1
            
            'set my range
            Set myRange = Range("K2:K" & rowNum)
                         
            'Use worksheet functions max and min to find values of percent change
            maxVal = Application.WorksheetFunction.Max(myRange)
            minVal = Application.WorksheetFunction.Min(myRange)
            
            'Reset my range
            Set myRange = Range("L2:L" & rowNum)
            
            'Use worksheet functions to find max value of total volume
            maxVolume = Application.WorksheetFunction.Max(myRange)
            
            'Add min and max values to the spreadsheet
            Range("Q2").Value = FormatPercent(maxVal, 2)
            Range("Q3").Value = FormatPercent(minVal, 2)
            Range("Q4").Value = maxVolume
    
            'Loop through range to find tickers for min and max values
            For j = 2 To rowNum
                If maxVal = Range("K" & j) Then
                    Range("P2").Value = Range("I" & j).Value
                End If
    
                If minVal = Range("K" & j) Then
                    Range("P3").Value = Range("I" & j).Value
                End If
    
                If maxVolume = Range("L" & j) Then
                    Range("P4").Value = Range("I" & j).Value
                End If
    
            Next j
        Debug.Print "Finished worksheet: " + ws.Name
                
        Next ws
End Sub
