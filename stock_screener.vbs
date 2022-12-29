Sub stock_screener():
    Dim lastRow As Integer
    Dim stock_ticker As String
    Dim open_price As Variant
    Dim close_price As Variant
    Dim price_change As Variant
    'Dim percentage_price_change As Double
    Dim volume_total As Integer
    Dim summary_table_row As Integer
    
    summary_table_row = 1
    
    Range("I" & summary_table_row).Value = "Ticker"
    Range("J" & summary_table_row).Value = "Yearly Change"
    Range("K" & summary_table_row).Value = "Percent Change"
    Range("L" & summary_table_row).Value = "Total Stock Volume"
    summary_table_row = summary_table_row + 1
    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    volume_total = 0
    
    
    For I = 2 To lastRow
        volume_total = volume_total + Cells(I, 7).Value
        open_price = Cells(I, 3).Value
            Range("J" & summary_table_row).Value = open_price - Cells(I, 4).Value
            Range("K" & summary_table_row).Value = (close_price - open_price) / open_price
            Range("L" & summary_table_row).Value = volume_total
            summary_table_row = summary_table_row + 1
            volume_total = 0
        End If
    Next I


End Sub