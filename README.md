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
        Dim total_stock As Double
        
        
        ' names for columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        
    
    
    
End Sub
