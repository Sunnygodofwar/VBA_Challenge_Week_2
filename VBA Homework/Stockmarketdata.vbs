Sub StockMarketData()

    ' Acticate worksheets
    
Dim WS As Worksheet

    For Each WS In ActiveWorkbook.Worksheets
    
    WS.Activate
    
    
        ' Determine the Last Row
        
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        ' Name Headers
        
        Cells(1, 9).Value = "Ticker"
        
        Cells(1, 10).Value = "Yearly Change"
        
        Cells(1, 11).Value = "Percent Change"
        
        Cells(1, 12).Value = "Total_Stock_Volume"
        
        Cells(1, 13).Value = "Ticker1"
        
        Cells(1, 14).Value = "Value"
        

        
        'Create Variable to hold Value
        
        Dim beginning_open_price As Double
        
        Dim ending_close_price As Double
        
        Dim Yearly_Change As Double
        
        Dim Ticker_Name As String
        
        Dim Percent_Change As Double
        
        Dim Volume As Double
        
        Volume = 0
        
        Dim Row As Double
        
        Row = 2
        
        Dim Column As Integer
        
        Column = 1
        
        Dim i As Long
        
        
        'Declare an initial opening price
        
        beginning_open_price = Cells(2, Column + 2).Value
        
        
         ' Loop through all ticker symbol
        
        For i = 2 To LastRow
        
         ' Check for changes in ticker
         
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
            
                ' Set Ticker name
                
                Ticker_Name = Cells(i, Column).Value
                
                Cells(Row, Column + 8).Value = Ticker_Name
                
                ' Set Close Price
                
                ending_close_price = Cells(i, Column + 5).Value
                
                ' Add Yearly Change
                
                Yearly_Change = ending_close_price - beginning_open_price
                
                Cells(Row, Column + 9).Value = Yearly_Change
                
                ' Add Percent Change
                
                If (beginning_open_price = 0 And ending_close_price = 0) Then
                
                    Percent_Change = 0
                    
                ElseIf (beginning_open_price = 0 And ending_close_price <> 0) Then
                
                    Percent_Change = 1
                    
                Else
                
                    Percent_Change = Yearly_Change / beginning_open_price
                    
                    Cells(Row, Column + 10).Value = Percent_Change
                    
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                    
                End If
                
                ' Add Total Volume
                
                Volume = Volume + Cells(i, Column + 6).Value
                
                Cells(Row, Column + 11).Value = Volume
                
                ' Add one to the summary table row
                
                Row = Row + 1
                
                ' reset the beginning open price
                
                beginning_open_price = Cells(i + 1, Column + 2)
                
                ' reset the Volume Total
                
                Volume = 0
                
            'if cells are the same ticker
            
            Else
            
                Volume = Volume + Cells(i, Column + 6).Value
                
            End If
            
        Next i
        
        ' Determine Last Row of Yearly Change per WS
        
        YCLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        
        ' Set color changes for positive = green; negative = red
        
        For j = 2 To YCLastRow
        
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
            
                Cells(j, Column + 9).Interior.ColorIndex = 4
                
            ElseIf Cells(j, Column + 9).Value < 0 Then
            
                Cells(j, Column + 9).Interior.ColorIndex = 3
                
            End If
            
        Next j
        
        Next WS
        
End Sub

