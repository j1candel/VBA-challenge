Attribute VB_Name = "Module1"
Sub stock_multi_year()

'Loop through each worksheet
For Each ws In Worksheets

    'Labeling the names of the columns
    ws.Cells(1, 9).Value = "Tick_Name"
    ws.Cells(1, 10).Value = "Yearly_Change"
    ws.Cells(1, 11).Value = "Percent_Change"
    ws.Cells(1, 12).Value = "VolumeStock_Change"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'Assigning each name a variable
    Dim tick_name As String
    Dim opening_total As Long
    Dim closing_total As Double
    Dim yearly_change As Double
    Dim percentage_change As Double
    Dim max As Double
    Dim min As Double
    Dim i As Long
    Dim volume_total As Double
    Dim greatest_total As Double
    Dim summary_table_row As Double

    'Setting closing_total, volume_total, & summary_table_row to a value
    closing_total = 0
    volume_total = 0
    summary_table_row = 2

    'Finding the last row
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Loop through all the tick_names
    For i = 2 To last_row

        'Check if we are within the same tick_name
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

            'Setting the tick_name, opening_total, closing_total, volume_total, & yearly change
            tick_name = ws.Cells(i, 1).Value
            opening_total = opening_total + ws.Cells(i, 3).Value
            closing_total = closing_total + ws.Cells(i, 6).Value
            volume_total = volume_total + ws.Cells(i, 7).Value
            yearly_change = closing_total - opening_total

                'Check if values are equal to 0
                If opening_total = 0 Then
                    yearly_change = 0
                Else
                    percentage_change = yearly_change / opening_total
                End If

            'Assigning ws.Cells a value
            ws.Range("I" & summary_table_row).Value = tick_name
            ws.Range("J" & summary_table_row).Value = yearly_change
            ws.Range("K" & summary_table_row).Value = percentage_change * 100
            ws.Range("L" & summary_table_row).Value = volume_total

            'Adding a row to summary_table_row
            summary_table_row = summary_table_row + 1

            'Resetting opening_total, closing_total, volume_total, yearly_change, & percentage_change
            opening_total = 0
            closing_total = 0
            volume_total = 0
            yearly_change = 0
            percentage_change = 0

    'If the cell immediately following a row is the same ticker_name
    Else

            'Adding a row to opening_total, closing_total, volume_total
            opening_total = opening_total + ws.Cells(i, 3).Value
            closing_total = closing_total + ws.Cells(i, 6).Value
            volume_total = volume_total + ws.Cells(i + 1, 7).Value
    End If

    Next i

    'Loop through yearly_change
    For i = 2 To summary_table_row

        'Check to see if yearly_change is positve
        If ws.Cells(i, 10).Value > 0 Then

            'If positive assign green
            ws.Cells(i, 10).Interior.ColorIndex = 4

        'Check to see if yearly_change is negative
        ElseIf ws.Cells(i, 10).Value < 0 Then

            'If negative assign red
            ws.Cells(i, 10).Interior.ColorIndex = 3

        End If

    Next i

        'Find the max value in percent_change
        max = Application.WorksheetFunction.max(ws.Range("K2:K" & summary_table_row))

    'Loop through percent_change
    For i = 2 To summary_table_row

        If max = ws.Cells(i, 11).Value Then

            'If equal to max value assign ws.Cells to ticker_name and max value
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(2, 17).Value = max

        Else

        End If

    Next i

        'Find the min value in percent_change
        min = Application.WorksheetFunction.min(ws.Range("K2:K" & summary_table_row))

    'Loop through percent_change
    For i = 2 To summary_table_row

        If min = ws.Cells(i, 11).Value Then

            'If equal to min value assign ws.Cells to ticker_name and max value
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(3, 17).Value = min

        Else

        End If

    Next i

    'Find the max value in volumestock_change
    greatest_total = Application.WorksheetFunction.max(ws.Range("L2:L" & summary_table_row))

        'Loop through volumestock_change
        For i = 2 To summary_table_row

        If greatest_total = ws.Cells(i, 12).Value Then

            'If equal to max volumestock_change value assign ws.Cells to volumestock_change and max value
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(4, 17).Value = greatest_total

        Else

        End If

    Next i

    'Change the format to decimals
    ws.Range("K2:K" & summary_table_row).NumberFormat = "0.00%"
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"

Next ws

End Sub



