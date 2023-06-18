Sub Stock_Volume():

    Dim Ticker As String
    
    Dim Yearly_Change As Double
    
    Dim Percent_Change As Double
    
    Dim open_price As Double
    
    Dim close_price As Double
    
    Dim start_price_row As Long
    start_price_row = 2
    
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
   
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    Dim ws As Worksheet
    
    lastRow = Range("A1").End(xlDown).Row

    For Each ws In ThisWorkbook.Worksheets

        For I = 2 To lastRow - 1
    
            If Cells(I, "A").Value <> Cells(I + 1, "A").Value Then
        
                Ticker = Cells(I, "A").Value
            
                Total_Stock_Volume = Total_Stock_Volume + Cells(I, "G").Value
            
                open_price = Cells(start_price_row, "C").Value
            
                close_price = Cells(I, "F").Value
                
                Yearly_Change = close_price - open_price
                
                Percent_Change = Yearly_Change / open_price
            
                Range("I" & Summary_Table_Row).Value = Ticker
            
                Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
            
                Range("J" & Summary_Table_Row).Value = Yearly_Change
            
                Range("J" & Summary_Table_Row).NumberFormat = "0.00"
            
                If Yearly_Change < 0 Then
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
                ElseIf Yearly_Change > 0 Then
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                 
                End If
            
            
                Range("K" & Summary_Table_Row).Value = Percent_Change
            
                Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
                Summary_Table_Row = Summary_Table_Row + 1
            
                start_price_row = I + 1
            
                Total_Stock_Volume = 0
            
            Else
        
                Total_Stock_Volume = Total_Stock_Volume + Cells(I, "G").Value
            
            End If
        
        Next I
        
    Next ws
    
    max_value = Application.WorksheetFunction.Max(Range("K2:K" & lastRow))
    
    Range("Q2").Value = max_value * 100 & "%"
    
    ticker_position = Application.WorksheetFunction.Match(max_value, Range("K2:K" & lastRow), 0)
    
    Range("P2").Value = Cells(ticker_position + 1, "I")
    
    min_value = Application.WorksheetFunction.Min(Range("K2:K" & lastRow))
    
    Range("Q3").Value = min_value * 100 & "%"
    
    ticker_position2 = Application.WorksheetFunction.Match(min_value, Range("K2:K" & lastRow), 0)
    
    Range("P3").Value = Cells(ticker_position2 + 1, "I")
    
    greatest_volume = Application.WorksheetFunction.Max(Range("L2:L" & lastRow))
    
    Range("Q4").Value = greatest_volume
    
    ticker_position3 = Application.WorksheetFunction.Match(greatest_volume, Range("L2:L" & lastRow), 0)
    
    Range("P4").Value = Cells(ticker_position3 + 1, "I")
    
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Yearly Change"
    Cells(1, "K").Value = "Percent Change"
    Cells(1, "L").Value = "Total Stock Volume"
    
    
            
End Sub

