# VBA-Challenge-Final
Sub Testing()
    ' to create loops through the worksheets
    Dim ws As Worksheet
    For Each ws In Worksheets
    
    ws.Activate
    
    
        ' to declare the variables
        Dim ticker As String
        Dim date_open As Double
        Dim date_close As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim total_stock_volume As Double
        Dim lastrow As Long
        Dim number_tickers As Integer
    

        'to set the start values
        date_open = 0
        date_close = 0
        yearly_change = 0
        percent_change = 0
        vol_total = 0
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        number_tickers = 0
        summary_table_row = 2
        

        
        ' names for columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        For i = 2 To lastrow
        
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                date_open = ws.Cells(i, 3).Value
            End If
            
            date_close = ws.Cells(i, 6).Value
                yearly_change = date_close - date_open
                ws.Cells(summary_table_row, 10).Value = yearly_change
            
            total_stock_volume = total_stock_volume + ws.Cells(i, 7)
             ws.Cells(summary_table_row, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(summary_table_row, 12).Value = vol_total
                
        If date_open = 0 And date_close = 0 Then
            percent_change = 0
            ws.Cells(summary_table_row, 11).Value = percent_change
            ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
        ElseIf date_open = 0 Then
            Dim new_stock As String
            new_stock = "New Stock"
            ws.Cells(summary_table_row, 11).Value = new_stock
        Else
            percent_change = yearly_change / date_open
            ws.Cells(summary_table_row, 11).Value = percent_change
            ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
            End If
                
            If yearly_change < 0 Then
            Cells(number_tickers + 1, 10).Interior.ColorIndex = 3
            ElseIf yearly_change > 0 Then
            Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
            Else
            Cells(number_tickers + 1, 10).Interior.ColorIndex = 6
            End If
            
                
        
        Next i
        
    Next ws
        
    
    
    
End Sub
    

