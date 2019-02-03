Sub StockMkt_easy()

        'Initial variable for holding the Ticker Symbols
    Dim Ticker_Symbol as String

        'Variables for holding the Total Volume of each ticker symbols
    Dim Volume_Total as Double
    Volume_Total = 0

        'Location for each Ticker Symbol, and Total Volume in the Summary Table
    Dim Summary_Table_Row as Integer
    Summary_Table_Row = 2


        'Adding values to the Summary Table.
    Range("K" & 1).Value = Ticker_Symbol
    Range("L" & 1).Value = Volume_Total


    For i = 2 To 70926

                    'Determine if we are still within the same ticker symbol, if not.
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value, Then

                    'Set the Ticker_Symbol
                Ticker_Symbol = cells(i, 1).Value

                     'Add the Volume_Total
                 Volume_Total = Volume_Total + cells(i, 7).Value

                    'Print the Ticker_Symbol in the Summary Table
                 Range("K" & Summary_Table_Row).value = Ticker_Symbol

                     'Print the Volume_Total in the Summary Table
                 Range("L" & Summary_Table_Row).value = Volume_Total

                    'Add one to the Summary Table Row
                Summary_Table_Row = Summary_Table_Row + 1

                    'Reset the Volume_Total
                Volume_Total = 0

                    'If the next cell down is the same Ticker Symbol as the previous one
            Else

                    'Add to the Volume Total
                Volume_Total = Volume_Total + cells(i, 7).Value 

            End If 

    Next i

End Sub

