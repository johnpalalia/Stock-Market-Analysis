Attribute VB_Name = "Module1"
Sub stock_analyzer():
    ' Iterate through each worksheet
    For Each ws In ActiveWorkbook.Worksheets

    ' Declare variables
    Dim lastRow As Long
    Dim ticker As String
    Dim Volume As Variant
    Dim yearOpen As Double
    Dim yearClose As Double
    Dim yearChange As Single
    Dim percentChange As Variant
    Dim rowCounter As Variant
    Dim resultsCounter As Variant
    
    ' Labels for the analysis fields
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"

    ' Results table starts at the second row
    resultsCounter = 2

    ' Counts the up to the last row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'ws.Cells(2, 13).Value = lastRow
    Volume = 0

    ' Capture the opening year value
    yearOpen = ws.Cells(2, 3).Value
    
    ' For loop to iterate through the second row to the last
    For rowCounter = 2 To lastRow
        
        'Sums the total volume
        Volume = Volume + ws.Cells(rowCounter, 7).Value

'-----MODERATE PART------
        
        ' The moderate part finds the percentChange, yearChange and conditionally formats the values under yearChange
        
        ' If the current ticker is the same as the one above and the ticker below is different than the current ticker then.
        If (ws.Cells(rowCounter - 1, 1).Value = ws.Cells(rowCounter, 1).Value And ws.Cells(rowCounter + 1, 1).Value <> ws.Cells(rowCounter, 1).Value) Then
           
           ' Capture the year close value at the current row
           yearClose = ws.Cells(rowCounter, 6).Value

           ' Calculate the change between close and open value
           yearChange = yearClose - yearOpen

           ' The if statement below handles a division by zero, the result can be changed to NaN stirng
           If yearOpen = 0 And yearClose <> 0 Then

                percentChange = yearClose / yearClose

           ElseIf yearOpen = 0 And yearClose = 0 Then

                percentChange = 0
            
           Else

                ' We calculate percent change, normally we multiply by 100 but we will add % and * 100 when displaying the value
                percentChange = (yearClose - yearOpen) / yearOpen

           End If

           ' We add one so that we move to the next year open
           yearOpen = ws.Cells(rowCounter + 1, 3).Value
           
            ' Display year change by ticker
            ws.Cells(resultsCounter, 10).Value = yearChange

            ' Conditional formating for year change: positive green, negative red
            If ws.Cells(resultsCounter, 10).Value < 0 Then

                ' Less than zero will be red
                ws.Cells(resultsCounter, 10).Interior.ColorIndex = 3
            Else

                ' Positive results and zero will be green
                ws.Cells(resultsCounter, 10).Interior.ColorIndex = 4
           
        End If
          
            ' Display percent change by ticker
           ws.Cells(resultsCounter, 11).Value = percentChange
           
            ' Format the numbers as a percentage with % and two decimal places
           ws.Cells(resultsCounter, 11).NumberFormat = "0.00%"
          
        End If

'-----EASY PART------

        'The easy part finds the total volume per ticker

        ' To display the total volume per ticker we check if the cell below is different
        If (ws.Cells(rowCounter + 1, 1).Value <> ws.Cells(rowCounter, 1).Value) Then
            
            ' Display tickers
            ws.Cells(resultsCounter, 9).Value = ws.Cells(rowCounter, 1).Value
            
            ' Print total volume in column 12
            ws.Cells(resultsCounter, 12).Value = Volume

            ' We reset the volume for the next ticker
            Volume = 0


           ' We increase the results counter by one for the next loop
           resultsCounter = resultsCounter + 1

        End If
       
    Next rowCounter
    
'-----HARD PART------

    ' Last part is the final results section for the greatest increase, greatest decrease, and greatest volume comparison

    ' Declare values to have our loop compare
    ws.Cells(2, 17).Value = 0
    ws.Cells(3, 17).Value = 0
    ws.Cells(4, 17).Value = 0

    ' Start our loop for the final results
    For resultsCounter = 2 To lastRow

        ' If the percent change is greater than the temporary value stored in cell (2,17)
        If ws.Cells(resultsCounter, 11).Value > ws.Cells(2, 17).Value Then

            ' Store the new temporary greater percent change value in cell (2,17)
            ws.Cells(2, 17).Value = ws.Cells(resultsCounter, 11).Value

            ' This statement will find the greatest value and print it to the results table
            ws.Cells(2, 16).Value = ws.Cells(resultsCounter, 9).Value
            
        End If

        ' If the percent change is lower than the temporary value stored in cell (3,17)
        If ws.Cells(resultsCounter, 11).Value < ws.Cells(3, 17).Value Then

            ' Store the new temporaty percent change value in cell (3,17)
            ws.Cells(3, 17).Value = ws.Cells(resultsCounter, 11).Value

            ' This statement will find the lowest value and print it the results table
            ws.Cells(3, 16).Value = ws.Cells(resultsCounter, 9).Value
        
        End If

        ' If the volume is greater than the temporary value in cell (4, 17)
         If ws.Cells(resultsCounter, 12).Value > ws.Cells(4, 17).Value Then

            ' Store the greater volume in cell (4,17)
            ws.Cells(4, 17).Value = ws.Cells(resultsCounter, 12).Value

            ' This statement will find the highest volume and print it in the results table
            ws.Cells(4, 16).Value = ws.Cells(resultsCounter, 9).Value
            
        End If
       
        Next resultsCounter
        
        'Format the results table
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"


    Next ws
 
End Sub
 
