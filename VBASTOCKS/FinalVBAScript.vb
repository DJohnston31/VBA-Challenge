Sub tickerdata()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    ' Declare Current as a worksheet object variable.
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets

        ' Add the word Ticker to the First Column of the summary table
        ws.Cells(1, 9).Value = "Ticker"

        ' Add the words Yearly Change to the Second Column of the summary table
        ws.Cells(1, 10).Value = "Yearly Change"

        ' Add the words Percent Change to the Third Column of the summary table
        ws.Cells(1, 11).Value = "Percent Change"

        ' Add the words Total Stock Volume  to the Fourth Column of the summary table
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Add the words Opening Stock Price to the Fifth Column of the summary table
        ws.Cells(1, 13).Value = "Opening Stock Price"

        ' Add the words Closing Stock Price to the Sixth Column of the summary table
        ws.Cells(1, 14).Value = "Closing Stock Price"

        ' Set an initial variable for holding the Ticker Symbol
        Dim Ticker As String

        ' Set an initial variable for opening price at beginning of the year
        Dim Opening_Price As Double
        
        ' Set an initial variable for closing price at end of the year
        Dim Closing_Price As Double

        ' Set an initial variable for Yearly change from opening price at the beginning of a given year to the closing price at the end of that year
        Dim Yearly_Change As Double

        ' Set an initial variable for % change from opening price at the beginning of a given year to the closing price at the end of that year
        Dim Percent_Change As Double

        ' Set an initial variable for Total Stock Volume
        Dim Volume As Double
        Volume = 0

        ' Keep track of the location for each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        ' Keep track of the RowCount of each ticker
        Dim RowCount As Integer
        RowCount = 1

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Loop through all stock transactions
        For i = 2 To LastRow

            ' Check if we are still within the same ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Set the Ticker
            Ticker = ws.Cells(i, 1).Value

            ' Print the Ticker in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker

            ' Set the Opening_Price
            Opening_Price = ws.Cells(i - RowCount + 1, 3).Value

            ' Print the Opening Price in the Summary Table
            ws.Range("M" & Summary_Table_Row).Value = Opening_Price
            
            ' Set the Closing_Price
            Closing_Price = ws.Cells(i, 6).Value

            ' Print the Closing Price in the Summary Table
            ws.Range("N" & Summary_Table_Row).Value = Closing_Price

            ' Set the Yearly_Change
            Yearly_Change = Closing_Price - Opening_Price

            ' Print the Yearly Change in the Summary Table
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

            'Check if the denominator value is zero before doing the calculation
            
            If Opening_Price = 0 Or IsEmpty(Opening_Price) Then
            ws.Range("K" & Summary_Table_Row).Value = "NA"
            
            Else
                        
            ' Set the Percent_Change
            Percent_Change = Yearly_Change / Opening_Price
            
            End If

            ' Print the Percent Change in the Summary Table
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

            ' Add to the Volume
            Volume = Volume + ws.Cells(i, 7).Value

            ' Print the Volume to the Summary Table
            ws.Range("L" & Summary_Table_Row).Value = Volume

            ' Reset the Volume
            Volume = 0

            ' Reset the RowCount
            RowCount = 0

            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1

            ' Reset the Opening_Price
            Opening_Price = 0

            ' Reset the Closing_Price
            Closing_Price = 0

            ' If the cell immediately following a row is the same ticker...
            Else

            ' Add to the Volume
            Volume = Volume + ws.Cells(i, 7).Value

            ' Add to the RowCount
            RowCount = RowCount + 1

            End If

        Next i

        ' Loop through all yearly change amounts for conditional formatting to highlight positive change in green and negative change in red.
        For i = 2 To LastRow

            ' Check if value is positive
            If ws.Cells(i, 10).Value > 0 Then

        ' Set the Cell Colors to Green

            ws.Cells(i, 10).Interior.ColorIndex = 4

        Else

        ' Set the Cell Colors to Red

            ws.Cells(i, 10).Interior.ColorIndex = 3

            End If

        Next i

        ' Add the words Greatest % Increase to the First Column of the highlights table
        ws.Cells(2, 17).Value = "Greatest % Increase"

        ' Add the words Greatest % Decrease to the First Column of the highlights table
        ws.Cells(3, 17).Value = "Greatest % Decrease"

        ' Add the words Greatest Total Volume to the Third Column of the summary table
        ws.Cells(4, 17).Value = "Greatest Total Volume"

        ' Add the word Ticker to the First Row of the highlights table
        ws.Cells(1, 18).Value = "Ticker"

        ' Add the word Value to the First Row of the highlights table
        ws.Cells(1, 19).Value = "Value"

        
        ' Set variable for Value with the Greatest % Increase Value
        Dim HighestPercentage As Double
        HighestPercentage = WorksheetFunction.Max(ws.Range("K:K"))
        
        ' Find Destination of cell with Greatest % Increase Value
        
        Set HighestPercentCell = ws.Range("K:K").Find(HighestPercentage, Lookat:=xlWhole)
        
        ' Insert values of the Greatest % Increase
    
        ws.Range("S2") = HighestPercentage
        ws.Range("S2").NumberFormat = "0.00%"
        'ws.Range("R2") = HighestPercentCell.Offset(, -3).Value
        
        ' Set variable for Value with the Greatest % Decrease Value
        Dim LowestPercentage As Double
        LowestPercentage = WorksheetFunction.Min(ws.Range("K:K"))
        
        ' Find Destination of cell with Greatest % Decrease Value
        
        Set LowestPercentageCell = ws.Range("K:K").Find(LowestPercentage, Lookat:=xlWhole)
        
        ' Insert values of the Greatest % Decrease
    
        ws.Range("S3") = LowestPercentage
        ws.Range("S3").NumberFormat = "0.00%"
        'ws.Range("R3") = LowestPercentageCell.Offset(, -3)
        
        ' Set variable for Value with the Greatest Total Volume
        Dim HighestVolume As Double
        HighestVolume = WorksheetFunction.Max(ws.Range("L:L"))

        ' Find Destination of cell with Greatest Total Volume

        Set HighestVolumeCell = ws.Range("L:L").Find(HighestVolume, Lookat:=xlWhole)
    
        ' Insert values of the Greatest Total Volume
    
        ws.Range("S4") = HighestVolume
        ws.Range("R4") = HighestVolumeCell.Offset(, -3)

    Next ws
        
End Sub