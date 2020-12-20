Attribute VB_Name = "Module1"
'loop to identify unique ticker symbols and insert them to column I
Sub stock_loop():
    'name columns
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Stock Volume"
    
    'declare variables
        Dim Ticker As String
        Dim Open_Rate As Double
        Dim Close_Rate As Double
        Dim Volume As Long
        Dim row_index As Integer
        Dim column_index As Integer
        Dim i As Long
        Dim last_row As Long
        
        i_summary = 2
        
        last_row = Cells(Rows.Count, 1).End(xlUp).Row
        
    ' MsgBox (last_row)
        For i = 2 To last_row:
            Ticker = Cells(i, 1).Value
            If Ticker <> Cells(i + 1, 1).Value Then
                Cells(i_summary, 9) = Ticker
                i_summary = i_summary + 1
        End If
    Next i
    
End Sub

'loop to populate yearly change, percent change and stock volume by unique ticker

Sub yearly_change():
    Dim row_index As Long
    Dim ticker_counter As Integer
    Dim ticker_index As Long
    Dim current_ticker As String
    Dim last_row As Long
    Dim last_unique_ticker_row As Long
    Dim start_row As Long
    Dim stock_volume As LongLong
        
    
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    last_unique_ticker_row = Cells(Rows.Count, 9).End(xlUp).Row
    ticker_counter = 0
    stock_volume = 0
    start_row = ticker_counter + 2
    
    For ticker_index = 2 To last_unique_ticker_row
        current_ticker = Cells(ticker_index, 9).Value
        For row_index = 2 To last_row
            If Cells(row_index, 1).Value = current_ticker Then
                ticker_counter = ticker_counter + 1
                stock_volume = stock_volume + Cells(row_index, 7)
            End If
        
        Next row_index
        
        Cells(ticker_index, 10).Value = Cells(start_row + ticker_counter - 1, 6).Value - Cells(start_row, 3).Value
        Cells(ticker_index, 11).Value = Cells(ticker_index, 10).Value / Cells(start_row, 3).Value * 100
        Cells(ticker_index, 12).Value = stock_volume
        start_row = start_row + ticker_counter
        ticker_counter = 0
        stock_volume = 0

    Next ticker_index

End Sub


