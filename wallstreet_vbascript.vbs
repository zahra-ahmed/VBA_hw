Sub stock_market():

' perform in every worksheet

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
    ws.Activate

    

' create summary table headers
    
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"


' variable for holding ticker name

    Dim ticker As String

' variable for open price
    Dim open_price As Double

' first open price
    open_price = Range("C2").Value

' variable for close price
    Dim close_price As Double


' initial variable for holding total stock volume

    Dim stock_volume As Double
    stock_volume = 0

' location of each ticker on summary table

    Dim summary_table_row As Integer
    summary_table_row = 2

' determine last row
    LR = Cells(Rows.Count, 1).End(xlUp).Row

' looping through all tickers

    For i = 2 To LR

        ' check if same ticker or not

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' set ticker name

            ticker = Cells(i, 1).Value

            ' print ticker name

            Range("I" & summary_table_row).Value = ticker

            ' set close price

            close_price = Cells(i, 6).Value

            ' calculate yearly change
            
            Range("J" & summary_table_row).Value = close_price - open_price

            ' yearly change color

            If Range("J" & summary_table_row).Value > 0 Then
                Range("J" & summary_table_row).Interior.ColorIndex = 4
            Else
                Range("J" & summary_table_row).Interior.ColorIndex = 3
            End If

            ' calculate percent change

            If open_price > 0 Then
                Range("K" & summary_table_row).Value = (close_price / open_price) - 1
                Range("K" & summary_table_row).NumberFormat = "0.00%"
            Else
                Range("K" & summary_table_row).Value = "NA"
            End If

            ' add to stock volume

            stock_volume = stock_volume + Cells(i, 7).Value


            ' print total stock volume

            Range("L" & summary_table_row).Value = stock_volume

            ' start next summary row

            summary_table_row = summary_table_row + 1

            ' reset stock volume

            stock_volume = 0

            ' next open price
            open_price = Cells(i + 1, 3).Value

          ' if same ticker

          Else

            ' add to the stock volume

            stock_volume = stock_volume + Cells(i, 7).Value

        End If

    Next i


    Next ws




MsgBox ("Complete")




End Sub






            





