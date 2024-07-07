Private Sub multiple_year_stock()
Dim information_row As Integer
Dim total_stock_volume, greatest_total_volume As LongLong
Dim ws As Worksheet
For Each ws In Worksheets


        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quaterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        Dim ticker, greatest_percentage_increase_ticker, greatest_percentage_decrease_ticker, greatest_total_volume_ticker As String
        Dim opening_price, closing_price, quaterly_change, percentage_change, greatest_percentage_increase, greatest_percentage_decrease, vEndRow As Double

'Initialize to zero
    opening_price = 0
    closing_price = 0
    quaterly_change = 0
    percentage_change = 0
    total_stock_volume = 0
    
'We start at row 2 because of the headers that are in place at row 1 of these columns
    information_row = 2
    
vEndRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
'Loop through all of the ticker information (choosing the longest of the three worksheets)
    For r = 2 To vEndRow
    
'Check if we are in the same ticker
        If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
            'Set the ticker, closing price, yearly change, and percentage change values
            ticker = ws.Cells(r, 1).Value
            closing_price = ws.Cells(r, 6).Value
            quaterly_change = closing_price - opening_price
            
            'Correct for the divide by zero error
            If opening_price = 0 Then
                percentage_change = quaterly_change
            ElseIf opening_price <> 0 Then
                percentage_change = quaterly_change / opening_price
            End If
            
            'Add to the total stock volume
            total_stock_volume = total_stock_volume + ws.Cells(r, 7).Value
            
            'Print (in order) the unique tickers, their yearly change, their percentage change, and their total stock volume to the information section of the worksheet
            ws.Range("I" & information_row).Value = ticker
            ws.Range("J" & information_row).Value = quaterly_change
            
            'Apply conditional formatting to the yearly change values
            If quaterly_change > 0 Then
                ws.Range("J" & information_row).Interior.ColorIndex = 4
            ElseIf quaterly_change < 0 Then
                ws.Range("J" & information_row).Interior.ColorIndex = 3
            End If
            
            'Format the percentage change values to display correctly
            ws.Range("K" & information_row).Value = percentage_change
            ws.Range("K" & information_row).NumberFormat = "0.00%"
            ws.Range("L" & information_row).Value = total_stock_volume
            
            
            'Add one to go to the next row of the information section
            information_row = information_row + 1
            
            'Reset the totals
            closing_price = 0
            quaterly_change = 0
            percentage_change = 0
            total_stock_volume = 0
            
        'If the cell immediately following the row is the same ticker...
        ElseIf Cells(r, 2).Value = "1/2/2022" Or Cells(r, 2).Value = "4/1/2022" Or Cells(r, 2).Value = "7/1/2022" Or Cells(r, 2).Value = "10/1/2022" Then
            opening_price = ws.Cells(r, 3).Value
        End If
            total_stock_volume = total_stock_volume + ws.Cells(r, 7).Value
        
           
  Next r
  ' Assign names for greatest increase,greatest decrease, and  greatest volume
        ws.Range("N1").Value = "Column Name"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker Name"
        ws.Range("P1").Value = "Value"
        
 


   'Go to the last row of column k

    kEndRow = ws.Cells(Rows.Count, "K").End(xlUp).Row

    'Define variable to initiate the second summery table value

    Increase = 0
    Decrease = 0
    Greatest = 0

        'find max/min for percentage change and the max volume Loop
        For k = 3 To kEndRow

            'Define previous increment to check
            last_k = k - 1

            'Define current row for percentage
            current_k = ws.Cells(k, 11).Value

            'Define Previous row for percentage
            prevous_k = ws.Cells(last_k, 11).Value

            'greatest total volume row
            volume = ws.Cells(k, 12).Value

            'Prevous greatest volume row
            prevous_vol = ws.Cells(last_k, 12).Value
            
            
            'Find the increase
            If Increase > current_k And Increase > prevous_k Then

                Increase = Increase

            ElseIf current_k > Increase And current_k > prevous_k Then

                Increase = current_k

                increase_name = ws.Cells(k, 9).Value

            ElseIf prevous_k > Increase And prevous_k > current_k Then

                Increase = prevous_k

                increase_name = ws.Cells(last_k, 9).Value

            End If

      
            'Find the decrease

            If Decrease < current_k And Decrease < prevous_k Then

                    Decrease = Decrease

            ElseIf current_k < Increase And current_k < prevous_k Then

                Decrease = current_k


                decrease_name = ws.Cells(k, 9).Value

            ElseIf prevous_k < Increase And prevous_k < current_k Then

                Decrease = prevous_k

                decrease_name = ws.Cells(last_k, 9).Value

            End If

       
           'Find the greatest volume

            If Greatest > volume And Greatest > prevous_vol Then

                Greatest = Greatest

                'define name for greatest volume
                'greatest_name = ws.Cells(k, 9).Value

            ElseIf volume > Greatest And volume > prevous_vol Then

                Greatest = volume

                'define name for greatest volume
                greatest_name = ws.Cells(k, 9).Value

            ElseIf prevous_vol > Greatest And prevous_vol > volume Then

                Greatest = prevous_vol

                'define name for greatest volume
                greatest_name = ws.Cells(last_k, 9).Value
                

            End If
            Next k
            
    'Get for greatest increase, greatest increase, and  greatest volume Ticker name
    
        ws.Range("O2").Value = increase_name
        ws.Range("O3").Value = decrease_name
        ws.Range("O4").Value = greatest_name
        ws.Range("P2").Value = Increase
        ws.Range("P3").Value = Decrease
        ws.Range("P4").Value = Greatest
        
    'Greatest increase and decrease in percentage format

    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"


  Next ws

End Sub

