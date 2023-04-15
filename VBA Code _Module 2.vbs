
Sub module2()
    
    For Each ws In Worksheets
       
    Dim Ticker As String
    Dim Row_Table As Double
    Row_Table = 2
    'print company's name starting at 2
    
    Dim Year_Change As Double
    Dim Closing_Price As Double
    Dim Open_Price As Double
    Dim Percent As Double
    Dim StockVolume As Double
    StockVolume = 0
    
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"
    
    ws.Cells(2, 15).Value = "Great % Increase"
    ws.Cells(3, 15).Value = "Great % Decrease"
    ws.Cells(4, 15).Value = "Great Total Volume"
    
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(1, 17).Value = "Ticker"
    
    
    Open_Price = ws.Cells(2, 3).Value
    LastRow_Main = ws.Cells(Rows.Count, 1).End(xlUp).Row
    LastRow_Table = ws.Cells(Rows.Count, 12).End(xlUp).Row
    
    For i = 2 To LastRow_Main 'starting at cell 2
    
             
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then 'starting in column 1 value of A
        
        'if the next cell does not equal to the current then execute
        Ticker = ws.Cells(i, 1).Value
        ws.Range("I" & Row_Table).Value = Ticker
        
        Closing_Price = ws.Cells(i, 6).Value
        'Column F has value of 6
        
        Year_Change = Closing_Price - Open_Price
                        
        'Column C has a value of 3
        ws.Range("J" & Row_Table).Value = Year_Change
            
        If Year_Change >= 0 Then
           ws.Range("J" & Row_Table).Interior.ColorIndex = 3
           'red code is 3 and green code is 4
        Else
            ws.Range("J" & Row_Table).Interior.ColorIndex = 4
        End If
    
        Percent = Year_Change / Open_Price
        
        ws.Range("K" & Row_Table).Value = Percent
        
         
        
        Open_Price = ws.Cells(i + 1, 3).Value
        
        ws.Range("L" & Row_Table).Value = StockVolume + ws.Cells(i, 7)
        '7 is the value for column G volume
        
        Row_Table = Row_Table + 1
        'what it used to be plus one
    
        
        Else
        
        StockVolume = StockVolume + ws.Cells(i, 7).Value
        'Total plus current
        
        
        End If
        
    
    Next i
        
        Dim LastRow_Results As Double
        
        Dim Greatest_Increase As Double
        Greatest_Increase = 0
        'inital value gives it a starting point
        Dim Greatest_IncreaseTicker As String
        
        Dim Greatest_Decrease As Double
        Greatest_Decrease = 0
        Dim Greatest_DecreaseTicker As String
        
        Dim Greatest_TotalVolume As Double
        Greatest_TotalVolume = 0
        Dim Greatest_TotalVolumeTicker As String

        
        'LastRow_Results = ws.Cells(Rows.Count, 11).End(x1Up).Row
    For i = 2 To LastRow_Main 'results from second table
        
        If ws.Range("K" & i).Value > Greatest_Increase Then
        Greatest_Increase = ws.Range("K" & i).Value
        'each value in column K is greater than greatest increase then set
        
        Greatest_IncreaseTicker = ws.Range("I" & i).Value
        
        ws.Range("P" & 2).Value = Greatest_Increase
        ws.Range("Q" & 2).Value = Greatest_IncreaseTicker
        
    End If
    Next i 'next value of i
        
       
    For i = 2 To LastRow_Main 'results from second table
        
        If ws.Range("K" & i).Value < Greatest_Decrease Then
        Greatest_Decrease = ws.Range("K" & i).Value
        'each value in column K is greater than greatest decrease then set
        
        Greatest_DecreaseTicker = ws.Range("I" & i).Value
                
        ws.Range("P" & 3).Value = Greatest_Decrease
        ws.Range("Q" & 3).Value = Greatest_DecreaseTicker
        
    End If
    Next i
        
        
    For i = 2 To LastRow_Main 'results from second table
        
        If ws.Range("L" & i).Value > Greatest_TotalVolume Then
        Greatest_TotalVolume = ws.Range("L" & i).Value
        'each value in column L is greater than greatest decrease then set
        
        
        Greatest_TotalVolumeTicker = ws.Range("I" & i).Value
        
        ws.Range("P" & 4).Value = Greatest_TotalVolume
        ws.Range("Q" & 4).Value = Greatest_TotalVolumeTicker
        
        End If
    Next i

Next ws


End Sub
