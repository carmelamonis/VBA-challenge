Attribute VB_Name = "Module1"
Sub doStock()

    'Doing it for all workseets
    For Each ws In Worksheets
                
        'Set an initial variable for holding the ticker
        Dim ticker As String
            
        'Dim volumeTotal As Long
        'volumeTotal = 0
        
        'Keep track of the location for each ticker in the summary table
        Dim sumTableRow As Long
        sumTableRow = 2
        
        'Gets the last row number
        Dim lastRow As Long
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                   
        'Create header table for the summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
                
        'Gets the opening price
        Dim openingPrice As Double
        
        'Define closing price
        Dim closingPrice As Double
        closingPrice = 0
        
        'Function definition of closing price & opening price
        Dim priceDifference As Double
        
        'Define percent change
        Dim percentChange As Double
        
        'Set an initial variable for holding total volume per ticker
        Dim volumeTotal As Double
        volumeTotal = 0
                   
        'Define the first starting price
        openingPrice = ws.Cells(2, 3).Value
        
        For i = 2 To lastRow
        
    
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Sets the ticker name
                ticker = ws.Cells(i, 1).Value
                
                'Print the ticker in the Summary table
                ws.Range("I" & sumTableRow).Value = ticker
                
                'Add the volume total
                volumeTotal = volumeTotal + ws.Cells(i, 7).Value
                'volumeTotal = CLng(volumeTotal) + ws.Cells(i, 7).Value
    
                
                'Gets the closing price
                closingPrice = ws.Cells(i, 6).Value
                'MsgBox ("Closing price " & closingPrice)
                
                'Print the priceDifference
                priceDifference = closingPrice - openingPrice
                ws.Range("J" & sumTableRow).Value = priceDifference
                'MsgBox ("Price difference " & priceDifference)
                
                'Changing the color of the yearly change cell
                'if the price difference is < 0, then turn red
                If priceDifference < 0 Then
                    ws.Range("J" & sumTableRow).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & sumTableRow).Interior.ColorIndex = 4
                End If
                
                
                'Print the percent change
                'percentChange = Round((priceDifference / openingPrice), 5)
                If openingPrice = 0 Then
                    percentChange = 0
                Else
                    percentChange = priceDifference / openingPrice
                End If
                ws.Range("K" & sumTableRow).Value = Format(percentChange, "0.00%")
                'MsgBox ("Percent change " & percentChange)
                
                
                'Print the total stock volume
                ws.Range("L" & sumTableRow).Value = volumeTotal
                'MsgBox ("Volume total " & volumeTotal)
                         
                'Assign the next opening price
                openingPrice = ws.Cells(i + 1, 3).Value
                
    
                'Add one to the summary table row
                sumTableRow = sumTableRow + 1
                
                'Reset volume total
                volumeTotal = 0
                'MsgBox ("Reset volume total " & volumeTotal)
           
                
        Else
                'Adds the volume total
                volumeTotal = volumeTotal + ws.Cells(i, 7).Value
                'volumeTotal = CLng((volumeTotal) + ws.Cells(i, 7).Value)
                'MsgBox (volumeTotal)
                
        End If
        
    Next i

'-------------Bonus Portion Creating New Table---------------------
        'Puts the labels for the new table
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
        'Gets the last row of the summary table
        Dim sumLastRow As Long
        sumLastRow = Cells(Rows.Count, 9).End(xlUp).Row
        
        'Define the greatest increase variable
        Dim perIncrease As Double
        perIncrease = 0
        
        'Define the greatest decrease variable
        Dim perDecrease As Double
        perDecrease = 0
        
        'Define greatest total volume
        Dim maxVol As Double
        maxVol = 0
        
        'Grabs the row number of the greatest increase, decrease, & volume
        Dim incRow, decRow, volRow As Long
        
        'Dim r As Long
        
        'Run through each value in the summary table
        For r = 2 To sumLastRow
            
            'For percentages greater or equal to 0
            If ws.Range("J" & r).Value >= 0 Then
            
                'if the current value is greater than current "greatest" value, replace it
                If ws.Range("K" & r).Value > perIncrease Then
                    perIncrease = ws.Range("K" & r).Value
                    incRow = r
                End If
            
            'if the percentage is less than 0
            Else
            
                'if the current value is less than the current "least" value, then replace it
                If ws.Range("K" & r).Value < perDecrease Then
                perDecrease = ws.Range("K" & r).Value
                decRow = r
                End If
            End If
            
            'if the current total vol is greater than the current "maximum" value, then replace it
            If ws.Range("L" & r).Value > maxVol Then
                maxVol = ws.Range("L" & r).Value
                volRow = r
            End If
                       
        Next r
        
        'Print out the statements in the chart
        'MsgBox (ws.Range("I" & incRow).Value)
        ws.Range("O2").Value = ws.Range("I" & incRow).Value
        ws.Range("P2").Value = Format(perIncrease, "0.00%")
        
        'MsgBox (ws.Range("I" & decRow).Value)
        ws.Range("O3").Value = ws.Range("I" & decRow).Value
        ws.Range("P3").Value = Format(perDecrease, "0.00%")
        
        'MsgBox (ws.Range("I" & volRow).Value)
        ws.Range("O4").Value = ws.Range("I" & volRow).Value
        ws.Range("P4").Value = maxVol
        
        'Fits everything in the column
        ws.Columns("A:P").AutoFit
    
    Next ws

End Sub

