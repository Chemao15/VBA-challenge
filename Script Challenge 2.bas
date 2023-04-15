Attribute VB_Name = "Module1"
Option Explicit
Sub stocks()
Dim name_stock As String
Dim yearly_change_amount As Double
Dim percent_chenge_amount As Double
Dim total_stock_vol As LongLong
Dim output_row As Long
Dim input_row As Long
Dim price_begin As Double
Dim price_final As Double
Dim lastrow As Long
Dim max_value_incre As Integer
Dim min_value_decre As Integer


lastrow = Cells(Rows.Count, 1).End(xlUp).Row
price_begin = Cells(2, 3).Value
total_stock_vol = 0
output_row = 2
percent_chenge_amount = 0
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"
Cells(2, 14).Value = "Greatest % increase"
Cells(3, 14).Value = "Greatest % decrease"
Cells(4, 14).Value = "Greatest total volume"
Cells(2, 15).Value = 0
Cells(2, 16).Value = 0
Cells(3, 15).Value = 0
Cells(3, 16).Value = 0
Cells(4, 15).Value = 0
Cells(4, 16).Value = 0




For input_row = 2 To lastrow
    name_stock = Cells(input_row, 1).Value
    total_stock_vol = total_stock_vol + Cells(input_row, 7).Value
    
    If Cells(input_row + 1, 1).Value <> name_stock Then
         'This is the las row ot the stock
        'Input
        price_final = Cells(input_row, 6).Value
        If total_stock_vol > Cells(4, 16).Value Then
           Cells(4, 16).Value = total_stock_vol
           Cells(4, 15).Value = name_stock
         End If
        'Calculations
        yearly_change_amount = price_final - price_begin
        percent_chenge_amount = yearly_change_amount / price_begin
        If percent_chenge_amount > Cells(2, 16).Value Then
           Cells(2, 16).Value = percent_chenge_amount
             Cells(2, 16).NumberFormat = "0.00%"
           Cells(2, 15).Value = name_stock
        End If
        If percent_chenge_amount < Cells(3, 16).Value Then
           Cells(3, 16).Value = percent_chenge_amount
             Cells(3, 16).NumberFormat = "0.00%"
           Cells(3, 15).Value = name_stock
        End If
        'Output
        Range("I" & output_row).Value = name_stock
        Range("J" & output_row).Value = yearly_change_amount
            If Cells(output_row, 10).Value > 0 Then
            Cells(output_row, 10).Interior.ColorIndex = 4
            Else
            Cells(output_row, 10).Interior.ColorIndex = 3
            End If
        Range("K" & output_row).Value = percent_chenge_amount
        Range("K" & output_row).NumberFormat = "0.00%"
        Range("L" & output_row).Value = total_stock_vol
           
       
       
        'Prepare for nex stock
        output_row = output_row + 1
        total_stock_vol = 0
        price_begin = Cells(input_row + 1, 3).Value
         
         
    End If
Next input_row



End Sub

