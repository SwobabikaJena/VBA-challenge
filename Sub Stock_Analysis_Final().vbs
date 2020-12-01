Sub Stock_Analysis_Final()

For Each ws In Worksheets

Sheets(ws.Name).Select

'---------------------------------------
'EASY AND MODERATE SOLUTION
'---------------------------------------

'Declare Variables
Dim lastrow As Long
Dim current_ticker As String
Dim next_ticker As String
Dim total_volume As Double
Dim summary_row As Integer
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim i As Double

'Assign Column Labels for Summary
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'set summary row initial value
summary_row = 2

'Set initial counter value for total stock volume
total_volume = 0

'Determine number of rows in the excel sheet
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Set open price for first ticker
opening_price = Cells(2, 3).Value


'Loop through all rows
For i = 2 To lastrow

'Define current and next ticker value
current_ticker = Cells(i, 1).Value
next_ticker = Cells(i + 1, 1).Value


'If next ticker is not same as current ticker then
If next_ticker <> current_ticker Then

'Adding current row's volume to toatal
total_volume = total_volume + Cells(i, 7).Value

'Assign values to summary columns
Cells(summary_row, 9).Value = current_ticker
Cells(summary_row, 12).Value = total_volume

'Define Closing price
closing_price = Cells(i, 6).Value

'Calculate change of closing price to opening price for that year
yearly_change = closing_price - opening_price

'Assign yearly change value to appropriate column
Cells(summary_row, 10).Value = yearly_change

'Conditional Formatting based on value
If yearly_change > 0 Then
Cells(summary_row, 10).Interior.ColorIndex = 4

ElseIf yearly_change < 0 Then

Cells(summary_row, 10).Interior.ColorIndex = 3
End If

'To handle division by 0 values

If opening_price <> 0 Then

'Calculate the percent change
percent_change = yearly_change / opening_price
Else
percent_change = 0

End If

'Assign value to percent change column and format it to percent
Cells(summary_row, 11).Value = percent_change
Cells(summary_row, 11).NumberFormat = "0.00%"

'Reset opening price for next ticker value
opening_price = Cells(i + 1, 3).Value

'Increment summary row value so that next ticker value can be put in that row
summary_row = summary_row + 1

'Reset Total stock volume so that it can be counted for the next ticker symbol
total_volume = 0

Else
total_volume = total_volume + Cells(i, 7).Value

End If

Next i

'---------------------------------------
'CHALLENGE SOLUTION:
'---------------------------------------

'--------------------------------------------------
'For Greatest % Increase
'----------------------------------------------------

'Assign names to summary table
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

'Declare variables
Dim j As Double
Dim max_percent As Double
Dim compare_percent As Double
Dim max_ticker As String

'Assign first value of the column as maximum percent and store corresponding ticker
max_percent = Cells(2, 11).Value
max_ticker = Cells(2, 9).Value


For j = 3 To lastrow

'Assign the next value in the column to another variable for comparison
compare_percent = Cells(j, 11).Value

'If next value is greater than equal to previous value then override max percent with new value
If compare_percent >= max_percent Then
max_percent = compare_percent
max_ticker = Cells(j, 9).Value


End If

Next j

'Assign the final values to cells in summary column
Cells(2, 17).Value = max_percent
Cells(2, 17).NumberFormat = "0.00%"
Cells(2, 16).Value = max_ticker

'--------------------------------------------------
'For Greatest % Decrease - Repeat above steps for percent decrease and maximum value
'--------------------------------------------------

Dim k As Double
Dim min_percent As Double
Dim compare_percent_1 As Double
Dim min_ticker As String

min_percent = Cells(2, 11).Value
min_ticker = Cells(2, 9).Value

For k = 3 To lastrow

compare_percent_1 = Cells(k, 11).Value

If compare_percent_1 <= min_percent Then
min_percent = compare_percent_1
min_ticker = Cells(k, 9).Value


End If

Next k

Cells(3, 17).Value = min_percent
Cells(3, 17).NumberFormat = "0.00%"
Cells(3, 16).Value = min_ticker

'--------------------------------------------------
'For Greatest Total Volume
'--------------------------------------------------

Dim l As Double
Dim max_volume As Double
Dim compare_volume As Double
Dim max_vol_ticker As String

max_volume = Cells(2, 12).Value
max_vol_ticker = Cells(2, 9).Value

For l = 3 To lastrow

compare_volume = Cells(l, 12).Value

If compare_volume >= max_volume Then
max_volume = compare_volume
max_vol_ticker = Cells(l, 9).Value


End If

Next l

Cells(4, 17).Value = max_volume
Cells(4, 16).Value = max_vol_ticker

Columns("I:Q").Select
Selection.EntireColumn.AutoFit


Next ws

End Sub

