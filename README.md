# VBA-Challenge-Final
Sub Testing()
    ' to create loops through the worksheets
    Dim ws As Worksheet
    For Each ws In Worksheets
    
        ' to declare the variables
        Dim ticker As String
        Dim date_open As Double
        Dim date_close As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim total_stock_volume As Double

        'to set the start values
        date_open = 0
        date_close = 0
        yearly_change = 0
        percent_change = 0
        total_stock_volume = 0
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        
        ' names for columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        For i = 2 To lastrow
        
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                date_open = ws.Cells(i, 3).Value
                
            End If
        
        
        
    Next ws
        
    
    
    
End Sub
