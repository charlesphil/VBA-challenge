Sub OutputInfo():
    ' Declare variables
    ' Variables to use in the information output area
    Dim info_row As Long
    Dim info_column As Integer
    Dim last_info_row As Long
    Dim year_change As Double
    ' Variables to use in the data space
    Dim first_data_row As Long
    Dim last_data_row As Long
    Dim year_open As Double
    Dim year_close As Double
    ' Variables to store maximum and minimum values for bonus area
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim max_volume As LongLong
    
    ' BONUS: For each worksheet, run the following
    For Each ws In Worksheets
        ' Set variable defaults
        info_row = 2
        ' Statement to get last row of a given sheet
        last_data_row = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        ' Assign header to first row, spanning four columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ' Output information below header in Columns I-L after looping through the data
        ' Loop through each row in the data and... (row denoted as r)
        For r = 2 To last_data_row
            ' If the ticker is the first ticker...
            If ws.Cells(r - 1, 1).Value <> ws.Cells(r, 1).Value Then
                ' Save the row of the first ticker to use later for finding total stock volume
                first_data_row = r
                ' Grab opening price of the year
                year_open = ws.Cells(r, 3).Value
            End If
            ' If the ticker is the last ticker...
            If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
                ' Output ticker to the information area
                ws.Cells(info_row, 9).Value = ws.Cells(r, 1).Value
                ' Grab closing price of the year
                year_close = ws.Cells(r, 6).Value
                ' Find the year change and output to information area
                year_change = year_close - year_open
                ws.Cells(info_row, 10).Value = year_change
                ' Find the percent change in decimal and output to information area if year_open price is not 0
                If year_open <> 0 Then
                    ws.Cells(info_row, 11).Value = year_change / year_open
                Else
                    ws.Cells(info_row, 11).Value = "NA"
                End If
                ' Get range of stock volume from the first occurance of the ticker to this last occurance of the ticker
                ws.Cells(info_row, 12).Value = WorksheetFunction.Sum(Range("G" & first_data_row & ":G" & r))
                ' Increment to the next row in the information area
                info_row = info_row + 1
            End If
        Next r
        
        ' Get last row in information area for use in conditional formatting
        last_info_row = ws.Cells(Rows.Count, 9).End(xlUp).row
        
        ' Conditional formatting for Yearly Change column
        ' Delete prior conditional formatting (just in case...)
        ws.Cells.FormatConditions.Delete
        ' If year change was greater than 0, make green
        ws.Range("J2:J" & last_info_row).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
        ws.Range("J2:J" & last_info_row).FormatConditions(1).Interior.Color = vbGreen
        ' If year change was less than 0, make red
        ws.Range("J2:J" & last_info_row).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        ws.Range("J2:J" & last_info_row).FormatConditions(2).Interior.Color = vbRed
        ' If year change was exactly 0, make gray
        ws.Range("J2:J" & last_info_row).FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0"
        ws.Range("J2:J" & last_info_row).FormatConditions(3).Interior.Color = RGB(128, 128, 128)
        ' Percentage formatting to two decimal places for Percent Change column
        ws.Range("K2:K" & last_info_row).NumberFormat = "0.00%"
        
        ' BONUS
        
        ' Headers and labels for bonus area
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ' Find greatest percent increase and decrease from column K and output into bonus area
        greatest_increase = WorksheetFunction.Max(ws.Range("K2:K" & last_info_row))
        greatest_decrease = WorksheetFunction.Min(ws.Range("K2:K" & last_info_row))
        ws.Cells(2, 17).Value = greatest_increase
        ws.Cells(3, 17).Value = greatest_decrease
        ' Format percentages
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        ' Find maximum stock volume from column L and output into bonus area
        max_volume = WorksheetFunction.Max(ws.Range("L2:L" & last_info_row))
        ws.Cells(4, 17).Value = max_volume
        
        ' Loop through each row in the outputted information and...
        For r = 2 To last_info_row
            ' If the current row's Percent Change is equal to the greatest increase, output ticker to P2
            If ws.Cells(r, 11).Value = greatest_increase Then
                ws.Cells(2, 16).Value = ws.Cells(r, 9).Value
            ' Otherwise, if it's equal to the greatest decrease, output ticker to P3
            ElseIf ws.Cells(r, 11).Value = greatest_decrease Then
                ws.Cells(3, 16).Value = ws.Cells(r, 9).Value
            End If
            ' Separate If block to check for stock volume, output ticket to P4 if match found
            If ws.Cells(r, 12).Value = max_volume Then
                ws.Cells(4, 16).Value = ws.Cells(r, 9).Value
            End If
        Next r
        
        ' Autofit all new columns (I through Q) in current sheet
        ws.Columns("I:Q").EntireColumn.AutoFit
    Next ws
End Sub