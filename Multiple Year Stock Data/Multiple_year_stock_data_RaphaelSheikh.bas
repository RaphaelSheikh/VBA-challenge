Attribute VB_Name = "Module1"
Sub StockAnalysis():

'Define variables

'Ticker

Dim Ticker As String

'Quarter Open

Dim QuarterOpen As Double

'Quarter Close

Dim QuarterClose As Double

'Quarterly Change

Dim QuarterlyChange As Double

'Total Stock Volume

Dim TotalStockVolume As Double

'Percent Change

Dim PercentChange As Double

'Set up a row to start

Dim StartRow As Integer

'The worksheet to execute the code in all worksheets at once in the workbook

Dim ws As Worksheet

'Loops through all the stocks for each quarter and outputs the following information:

For Each ws In Worksheets

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    StartRow = 2
    previous_i = 1
    TotalStockVolume = 0
    
    EndRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
        'For each Ticker loop the quarterly change, percent change, and total stock volume

        For i = 2 To EndRow
        
            'If Ticker change or not equal to the previous one excute to record
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Get the Ticker
            
            Ticker = ws.Cells(i, 1).Value
            
            'Next Ticker

            previous_i = previous_i + 1
            
            'Value of first day open and last day close

            QuarterOpen = ws.Cells(previous_i, 3).Value
            QuarterClose = ws.Cells(i, 6).Value

            'For loop to sum the total stock volume using volume

            For j = previous_i To i
                
                TotalStockVolume = TotalStockVolume + ws.Cells(j, 7).Value
                
            Next j
            
            If QuarterOpen = 0 Then
            
                PercentChange = QuarterClose
                
            Else
                QuarterlyChange = QuarterClose - QuarterOpen
                PercentChange = QuarterlyChange / QuarterOpen
                
            End If
            
            'Get values in the worksheet summary table

            ws.Cells(StartRow, 9).Value = Ticker
            ws.Cells(StartRow, 10).Value = QuarterlyChange
            ws.Cells(StartRow, 11).Value = PercentChange
            ws.Cells(StartRow, 11).NumberFormat = "0.00%"
            ws.Cells(StartRow, 12).Value = TotalStockVolume
            
            StartRow = StartRow + 1
            
            TotalStockVolume = 0
            QuarterlyChange = 0
            PercentChange = 0
            
            previous_i = i
        
        End If
        
'Loop is completed

    Next i
    
'The second summery table

    'Go to the last row of column k

    kEndRow = ws.Cells(Rows.Count, "K").End(xlUp).Row

    'Define variable to initiate the second summery table value

    Increase = 0
    Decrease = 0
    Greatest = 0

        'Find max/min for percentage change and the max volume loop
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

                'define name for increase percentage
                'increase_name = ws.Cells(k, 9).Value

            ElseIf current_k > Increase And current_k > prevous_k Then

                Increase = current_k

                'define name for increase percentage
                increase_name = ws.Cells(k, 9).Value

            ElseIf prevous_k > Increase And prevous_k > current_k Then

                Increase = prevous_k

                'define name for increase percentage
                increase_name = ws.Cells(last_k, 9).Value

            End If

            'Find the decrease

            If Decrease < current_k And Decrease < prevous_k Then

                'Define decrease as decrease

                Decrease = Decrease

                'Define name for increase percentage

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

    'Assign names for greatest increase, greatest decrease, and greatest volume

    ws.Range("N1").Value = " "
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"

    'Get for greatest increase, greatest increase, and greatest volume Ticker name
    ws.Range("O2").Value = increase_name
    ws.Range("O3").Value = decrease_name
    ws.Range("O4").Value = greatest_name
    ws.Range("P2").Value = Increase
    ws.Range("P3").Value = Decrease
    ws.Range("P4").Value = Greatest

    'Greatest increase and decrease in percentage format

    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"

' Conditional formatting columns colors

'The end row for column J

    jEndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row


        For j = 2 To jEndRow

            'if greater than or less than zero
            If ws.Cells(j, 10) > 0 Then

                ws.Cells(j, 10).Interior.ColorIndex = 4

            Else

                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If

        Next j

'Excute to next worksheet
Next ws

End Sub
