Sub stock_mkt()

    'procedures for all sheets'
    For Each ws in Worksheets

        'Find last row and set variable'
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

        'Set variable to hold name for ticker symbol'
        Dim Ticker_Symbol As String
        Ticker_Symbol = " "

        'Set variable for yearly change'
        Dim yearly_change As Double
        yearly_change = 0

        'Set variable for opening price'
        Dim opening_price As Double
        opening_price = 0

        'Set variable for closing price'
        Dim closing_price As Double
        closing_price = 0

        'Set variable for percent change'
        Dim percent_change As Double
        percent_change = 0

        'Set variable to hold total volume for each ticker symbol'
        Dim total_stock_volume As variant
        total_stock_volume = 0

        'Keep track of location of each Stock in new table'
        Dim Stock_Row As Long
        Stock_Row = 2

        'Setting opeing price to first open value, which is (2, 3)'
        opening_price = ws.Cells(2,3).Value

        'Add column headers for Ticker, Yearly Change, Percent Change, Total Stock Volume'
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Stock Volume"

        'creating loop'
        For i = 2 to LastRow

            'If next value in loop does not equal predecessor (Then)'
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                'moving on to next ticker symbol value'
                Ticker_Symbol = ws.Cells(i, 1).Value

                ' " " '
                closing_price = ws.Cells(i, 6).Value

                'calculate yearly change'
                yearly_change = closing_price - opening_price

                'new If statement to calculate percent change'
                If opening_price <> 0 Then

                percent_change = (yearly_change / opening_price) * 100

                End If

                'add values of variables to new column for Ticker Symbol, yearly change, percent change, and total stock volume'
                ws.Range("I" & Stock_Row).Value = Ticker_Symbol

                ws.Range("J" & Stock_Row).Value = yearly_change

                'adding c string to make percent change into string to display %'
                ws.Range("K" & Stock_Row).Value = (Cstr(percent_change) & "%")

                ws.Range("L" & Stock_Row).Value = total_stock_volume

                'setting conditional formatting highlighting positive and negative change'
                If (yearly_change > 0) Then

                ws.Range("J" & Stock_Row).Interior.ColorIndex = 4

                Else

                ws.Range("J" & Stock_Row).Interior.ColorIndex = 3

                End If

                'moving stock row and resetting variables for opening price, price change, and total volume'
                Stock_Row = Stock_Row + 1

                opening_price = ws.Cells(i + 1, 3).Value

                percent_change = 0

                total_stock_volume = 0

            Else

                'if next value does equal predecessor, then total stock volume to previous total'
                total_stock_volume = total_stock_volume + ws.Cells(i + 1, 7).Value

            End If

        Next i 

    Next ws

End Sub