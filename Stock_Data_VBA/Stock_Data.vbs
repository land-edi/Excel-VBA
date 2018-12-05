Sub StockAnalysis()

    ' --------------------------------------------
    ' ITERATE THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets

        ' -------------------------------------------------------------
        ' Provide the summary for each stock ticker and find greatests
        ' --------------------------------------------------------------

        'Add Column headers for the summary table
        ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
        
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        ' Define a variable to hold the Ticker name
        Dim Ticker_Name As String

        ' Define a variable to hold the total volume
        Dim Ticker_Volume As Double
        Ticker_Volume = 0
        
        'Define a variable to hold the yearly open value
        Dim Ticker_Open As Double
        Ticker_Open = 0
        
        'Define a variable to hold the percent change
        Dim Pct_Change As Double
        Pct_Change = 0

        ' Keep track of the location for each Ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'Add column headers and row names for the Greatest table
        ws.Range("P1:Q1").Value = Array("Ticker", "Value")
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Define variables to hold the greatest values
        Dim Ticker_Increase As String
        Dim Ticket_Decrease As String
        Dim Ticker_Total As String
        Dim Great_Increase As Double
        Dim Great_Decrease As Double
        Dim Great_Total As Double
        Great_Increase = 0
        Great_Decrease = 0
        Great_Total = 0

        ' Iterate through all Tickers
        For i = 2 To LastRow

            ' Check if we are at the first record of the year
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

                ' Set the open value
                Ticker_Open = ws.Cells(i, 3).Value
                
            End If
            
            ' Check if we are at the last record of the ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

                ' Set the Ticker name
                Ticker_Name = ws.Cells(i, 1).Value

                ' Add to the Ticker Total volume
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value

                ' Print the Ticker name in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

                ' Print the Yearly change to the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = ws.Cells(i, 3).Value - Ticker_Open
				If ws.Cells(i, 3).Value <= Ticker_Open Then
					ws.Range("J" & Summary_Table_Row).Interior.Color = vbRed
				Else
					ws.Range("J" & Summary_Table_Row).Interior.Color = vbGreen
				End If
                
                ' Print the percent change to the Summary Table
                If Ticker_Open <> 0 Then
                    Pct_Change = (ws.Cells(i, 3).Value - Ticker_Open) / Ticker_Open
                    ws.Range("K" & Summary_Table_Row).Value = Pct_Change
                    ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                End If
                    
                ' Print the Total Volume to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Ticker_Volume

                '-------------------------------------------------------------
                'Check for greatest values
                '-------------------------------------------------------------
                
                'Check for Greatest increase
                If Pct_Change >= Great_Increase Then
                    Great_Increase = Pct_Change
                    Ticker_Increase = Ticker_Name
                End If
                
                'Check for Greatest decrease
                If Pct_Change <= Great_Decrease Then
                    Great_Decrease = Pct_Change
                    Ticker_Decrease = Ticker_Name
                End If
                
                'Check for Greatest Volume
                If Ticker_Volume >= Great_Total Then
                    Great_Total = Ticker_Volume
                    Ticker_Total = Ticker_Name
                End If


                '----------------------------------
                'Reset values for next Ticker
                '----------------------------------
                
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
      
                ' Reset the Ticker Total volume
                Ticker_Volume = 0

            ' If the cell immediately following a row is the same brand...
            Else

                ' Add to the Ticker Total volume
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value

            End If
            
        Next i

        '-------------------------------------
        'Print the Greatest Values
        '-------------------------------------
        ws.Range("P2").Value = Ticker_Increase
        ws.Range("P3").Value = Ticker_Decrease
        ws.Range("P4").Value = Ticker_Total
            
        ws.Range("Q2").Value = Great_Increase
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").Value = Great_Decrease
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").Value = Great_Total

    Next ws

End Sub
