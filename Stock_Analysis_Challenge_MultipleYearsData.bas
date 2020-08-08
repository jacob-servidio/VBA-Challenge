Attribute VB_Name = "Module1"
Sub Stock_Analysis_Challenge_MultipleYearsData()
'Create a loop through all stocks for one year
'Output ticker symbol
'YTD change formatted as green(+) or red(-)
'YTD% change
'Total trading volume
'---------------------------------------------------

'loop through worksheets
Dim ws As Worksheet
For Each ws In Worksheets

'summary table headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Volume"

'declare variables

Dim ticker As String
Dim volume_total As Double
volume_total = 0
Dim year_open As Double
Dim year_close As Double
Dim percent_change As Double
Dim summary_table_row As Long
summary_table_row = 2
Dim lastrow As Long
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'loop through data
For i = 2 To lastrow

If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
    
    year_open = ws.Cells(i, 3).Value
    
End If

'grab volume from each entry to form volume_total
volume_total = volume_total + ws.Cells(i, 7)

'if ticker changed
If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

'move needed values to Summary table
ws.Cells(summary_table_row, 9).Value = ws.Cells(i, 1).Value
ws.Cells(summary_table_row, 12).Value = volume_total
year_close = ws.Cells(i, 6).Value
year_change = year_close - year_open
ws.Cells(summary_table_row, 10).Value = year_change

'highlight yearly change green(+), red(-)
If year_change >= 0 Then
    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
Else
    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
End If

'calculate % change and format as %
If year_open = 0 And year_close = 0 Then
        percent_change = 0
        ws.Cells(summary_table_row, 11).Value = percent_change
        ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
        
ElseIf year_open = 0 Then
    Dim percent_change_NewStock As String
    percent_change_NewStock = "New Stock"
    ws.Cells(summary_table_row, 11).Value = percent_change_NewStock
Else
    percent_change = (year_close - year_open) / year_open
        ws.Cells(summary_table_row, 11).Value = percent_change
        ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
        
End If

'add row to summary table
summary_table_row = summary_table_row + 1

'reset values to 0
volume_total = 0
year_open = 0
year_close = 0
year_change = 0
percent_change = 0

End If

Next i

'bonus section
'----------------------------------

'titles for best/worst performance table

ws.Cells(2, 15).Value = "greatest % inc"
ws.Cells(3, 15).Value = "greatest % dec"
ws.Cells(4, 15).Value = "greatest total vol"
ws.Cells(1, 16).Value = "ticker"
ws.Cells(1, 17).Value = "value"


'assign lastrow for summary table
lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
'declare variables
Dim best_stock As String
Dim best_value As Variant
best_value = ws.Cells(2, 11).Value
Dim worst_stock As String
Dim worst_value As Double
worst_value = ws.Cells(2, 11).Value
Dim most_volStock As String
Dim most_volValue As Double
most_volValue = ws.Cells(2, 12).Value

'loop through summary table
For j = 2 To lastrow

'determine best performer
If ws.Cells(j, 11).Value > best_value And ws.Cells(j, 11).Value <> "New Stock" Then
    best_value = ws.Cells(j, 11).Value
    best_stock = ws.Cells(j, 9).Value

End If

'determine worst performer
If ws.Cells(j, 11).Value < worst_value Then
    worst_value = ws.Cells(j, 11).Value
    worst_stock = ws.Cells(j, 9).Value
    
End If

'determin greatest total volume traded
If ws.Cells(j, 12).Value > most_volValue Then
    most_volValue = ws.Cells(j, 12).Value
    most_volStock = ws.Cells(j, 9).Value

End If


'send values to table
ws.Cells(2, 16).Value = best_stock
ws.Cells(2, 17).Value = best_value
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 16).Value = worst_stock
ws.Cells(3, 17).Value = worst_value
ws.Cells(3, 17).NumberFormat = "0.00%"
ws.Cells(4, 16).Value = most_volStock
ws.Cells(4, 17).Value = most_volValue

Next j

Next ws


End Sub
