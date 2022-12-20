Sub tickers()
    For Each ws In Worksheets

        'Creating all the variables I will need
        Dim lastRow As Long
        Dim ticker As String
        Dim openPrice As Double
        Dim closingPrice As Double
        Dim percentChange As Double
        Dim totalVolume As Double
        Dim summaryRow As Integer
        Dim totalChange As Double

        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatestVolume As Double
        Dim incTick As String
        Dim decTick As String
        Dim volTick As String


        'Setting the formatting for the columns we will be creating that need formatting
        ws.Columns("J").NumberFormat = "##.##"
        ws.Columns("K").NumberFormat = "0.00%"
        
        'Naming the new columns and rows 
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"


        'Finding the last row of data and saving it
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Initializing all variables to be default values before looping through a worksheets data
        ticker = ""
        openPrice = ws.Cells(2, 3).Value
        closingPrice = 0
        percentChange = 0
        totalVolume = 0
        summaryRow = 2

        greatestDecrease = 0
        greatestIncrease = 0
        greatestVolume = 0
        incTick = ""
        decTick = ""
        volTick = ""
        
        'Looking through a worksheet from the first value row to the last row with data
        For i = 2 To lastRow
            
            'Checking if the ticker values in the current cell and the next cell are different to know if the ticker data is ending
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                'Setting the ticker value to be the current ticker
                ticker = ws.Cells(i, 1).Value
                'Setting the closing price to the closing price of the final row of data
                closingPrice = ws.Cells(i, 6)
                'Adding up the total volume with the final line
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                'Calculating the total and percentage change from the opening price at the beginning of the year and the closing price at the end of the year
                totalChange = closingPrice - openPrice
                percentChange = totalChange / openPrice

                ' Print the ticker in the Summary Table
                ws.Range("I" & summaryRow).Value = ticker

                ' Print the total volume to the Summary Table
                ws.Range("L" & summaryRow).Value = totalVolume
                
                
                ' Print the Yearly Change
                ws.Range("J" & summaryRow).Value = totalChange
                'Checking if the total change is positive or negative and coloring the box appropriately
                If (totalChange > 0) Then
                    ws.Range("J" & summaryRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & summaryRow).Interior.ColorIndex = 3
                End If

                ' Print the percentage change
                ws.Range("K" & summaryRow).Value = percentChange

                'Checking if the current percent change is the greatest or the lowest, and then if the volume is the greatest
                If (greatestIncrease < percentChange) Then
                    greatestIncrease = percentChange
                    incTick = ticker
                End If
                If (greatestDecrease > percentChange) Then
                    greatestDecrease = percentChange
                    decTick = ticker
                End If
                If (greatestVolume < totalVolume) Then
                    greatestVolume = totalVolume
                    volTick = ticker
                End If

                ' Add one to the summary table row
                summaryRow = summaryRow + 1
                
                ' Reset the volume Total
                totalVolume = 0
                openPrice = ws.Cells(i + 1, 3).Value

            ' If the cell immediately following a row is the same brand...
            Else

                ' Add to the Total Volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value

            End If

        Next i
        'Filling in the greatest increase, decrease and volume parts of the table. 
        ws.Range("P2").Value = greatestIncrease
        ws.Range("O2").Value = incTick
        ws.Range("P2").NumberFormat = "0.00%"

        ws.Range("P3").Value = greatestDecrease
        ws.Range("O3").Value = decTick
        ws.Range("P3").NumberFormat = "0.00%"

        ws.Range("P4").Value = greatestVolume
        ws.Range("O4").Value = volTick
        ws.Range("P4").NumberFormat = "General"



    Next ws
End Sub





