Sub StockMkt_Mod()
    
        'Create Intial variables
    Dim total as Double
    Dim next_ticker as Integer
    Dim ticker as String
    Dim ws as Worksheet
    Dim openprice as Double
    Dim closeprice as Double
    Dim change as Double
    Dim lastrow as Double

        'add values to the created variables
    total = 0
    next_ticker = 2

            'add the opening stock price to variable
            openprice = ws.Cells(i, 3).Value

    For i = 2 To ws.Cells(Rows.Count, 1).End(xlp).Row

        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then

            total = total + ws.Cells(i, 7).Value
        
        Else
            ticker = ws.Cells(i, 1).Value
            total = total + ws.Cells(i, 7).Value
            ws.Cells(next_ticker, 9).Value = ticker
            ws,Cells(next_ticker, 12).Value = total
            total = 0

            'add the closing price to the variable
            closeprice = ws.Cells(i, 6).Value  

                'change of stock price at the end of the year compared to the beginning of the year, and cell formatting
            ws.Cells(next_ticker, 10).Value = closeprice - openprice
            ws.Cells(next_ticker, 11).NumberFormat = "00.000000000"
            change = ws.Cells(next_ticker, 10).Value

                'change the color depending if the change is positive or negative
            If change > 0 Then
            ws.Cells(next_ticker, 10).Interior.ColorIndex = 4
            Else 
            ws.Cells(next_ticker, 10).Interior.ColorIndex = 3
            End If

                'precent change with cell formatting
            If change = 0 Then
            ws.Cells(next_ticker, 11) = 0
            Elseif openprice = 0 Then
            ws.Cells(next_ticker, 11).Value = ""
            Else
            ws.Cells(next_ticker, 11).Value = change / openprice
            ws.Cells(next_ticker, 11).NumberFormat = "0.00%"
            End if

                'open price for the next stock
            openprice = ws.Cells(i + 1, 3).Value
            next_ticker = next_ticker + 1

        End if

        Next i 

                'add header to identify the new columns
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"

End Sub