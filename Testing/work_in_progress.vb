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
    
    ' Set variable defaults
    info_row = 2
    data_column = 1
    ' Statements to get last row and column of a given sheet
    last_data_row = Cells(Rows.Count, 1).End(xlUp).row
    
    ' Assign header to first row, spanning four columns
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    ' Output information below header in Columns I-L after looping through the data
    ' Loop through each row in the data and... (row denoted as r)
    For r = 2 To last_data_row
        ' If the ticker is the first ticker...
        If Cells(r - 1, 1).Value <> Cells(r, data_column).Value Then
            ' Save the row of the first ticker to use later for finding total stock volume
            first_data_row = r
            ' Grab opening price of the year
            year_open = Cells(r, 3).Value
        End If
        ' If the ticker is the last ticker...
        If Cells(r + 1, 1).Value <> Cells(r, data_column).Value Then
            ' Output ticker to the information area
            Cells(info_row, 9).Value = Cells(r, data_column).Value
            ' Grab closing price of the year
            year_close = Cells(r, 6).Value
            ' Find the year change and output to information area
            year_change = year_open - year_close
            Cells(info_row, 10).Value = year_change
            ' Find the percent change in decimal and output to information area
            Cells(info_row, 11).Value = year_change / year_open
            ' Get range of stock volume from the first occurance of the ticker to this last occurance of the ticker
            Cells(info_row, 12).Value = WorksheetFunction.Sum(Range("G" & first_data_row & ":G" & r))
            ' Increment to the next row in the information area
            info_row = info_row + 1
        End If
    Next r
    
    ' Get last row in information area for use in conditional formatting
    last_info_row = Cells(Rows.Count, 9).End(xlUp).row
    
    ' Conditional formatting for Yearly Change column
    ' If year change was greater than 0, make green
    Range("J2:J" & last_info_row).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
    Range("J2:J" & last_info_row).FormatConditions(1).Interior.Color = vbGreen
    ' If year change was less than 0, make red
    Range("J2:J" & last_info_row).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    Range("J2:J" & last_info_row).FormatConditions(2).Interior.Color = vbRed
    ' If year change was exactly 0, make gray
    Range("J2:J" & last_info_row).FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0"
    Range("J2:J" & last_info_row).FormatConditions(3).Interior.Color = RGB(128, 128, 128)
    
    ' Percentage formatting to two decimal places for Percent Change column
    Range("K2:K" & last_info_row).NumberFormat = "0.00%"
    
    ' BONUS
    
    ' Headers and labels for bonus area
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    
End Sub