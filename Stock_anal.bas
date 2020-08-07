Attribute VB_Name = "Module1"
Sub Stock_anal()

'To make the tables I need the headers as an array, but making arrays manually
'in VBA is super annoying so I made them as strings and split them into arrays

Dim Titles1 As String
Dim Titles2 As String
Dim Titles3 As String
Dim Sheet_labels As String
Titles1 = "Ticker-Yearly Change-Percent Change-Total Stock Volume"
Titles2 = "Ticker-Value"
Titles3 = "Greatest % Increase-Greatest % Decrease-Greeatest Total Volume"
Sheet_labels = "2014-2015-2016"

Dim table_1_titles() As String
Dim table_2_titles() As String
Dim table_2_row_titles() As String
Dim sheet_names() As String

table_1_titles = Split(Titles1, "-")
table_2_titles = Split(Titles2, "-")
table_2_rows = Split(Titles3, "-")
sheet_names = Split(Sheet_labels, "-")

Dim i As Long
Dim row_count As Long


'This will initiate the larger loop which will run across worksheets

For j = 1 To 3

If j = 1 Then
Sheets("2014").Activate
ElseIf j = 2 Then
Sheets("2015").Activate
ElseIf j = 3 Then
Sheets("2016").Activate
End If

'The following function will construct the output tables.

'This will construct the first table:

For i = 0 To 3
Cells(1, (i + 9)) = table_1_titles(i)
Next i

'This will construct the second table
For i = 0 To 1
Cells(1, (i + 16)) = table_2_titles(i)
Next i

For i = 0 To 2
Cells((2 + i), 15) = table_2_rows(i)
Next i


'Well need to know how many rows there are. There might be an easier way to do this
'but meh


row_count = 0

Do While Cells(row_count + 2, 1) <> ""
row_count = row_count + 1
Loop

'adding 5 to row_count to for some overkill protection against off-by-one errors
row_count = row_count + 5

'This will populate the first table
Dim current_symbol As String
Dim first_price As Double
Dim last_price As Double
Dim counter As Long


counter = 2
Cells(counter, 12) = 0
current_symbol = Cells(2, 1)
first_price = Cells(2, 3)

For i = 2 To row_count
Cells(counter, 12) = Cells(counter, 12) + Cells(i, 7)

'This will trigger on the last row of a ticker symbol.  It will populate a row of Table1
If Cells(i + 1, 1) <> Cells(i, 1) Then

Cells(counter + 1, 12) = 0
last_price = Cells(i, 6)
Cells(counter, 9) = current_symbol
Cells(counter, 10) = last_price - first_price
'checks if percentage increase would drop a div/0 error, and if so inputs 0
If first_price = 0 Then
Cells(counter, 11) = 0
Else
Cells(counter, 11) = Cells(counter, 10) / first_price
End If


'This will reset the values of the variables before the end of IF; current_symbol and
'first price will need to look one row ahead.  Last_price is set to 0 for error visiblity

current_symbol = Cells(i + 1, 1)
first_price = Cells(i + 1, 3)
last_price = 0
counter = counter + 1
End If
Next i

'This will populate the next table
Dim gainest_ticker As String
Dim gainest_value As Double
Dim lossest_ticker As String
Dim lossest_value As Double
Dim totalest_ticker As String
Dim totalest_value As Double

gainest_value = 0
lossest_value = 0
totalest_value = 0

For i = 2 To row_count

If Cells(i, 11) > gainest_value Then
gainest_value = Cells(i, 11)
gainest_ticker = Cells(i, 9)

ElseIf Cells(i, 11) < lossest_value Then
lossest_value = Cells(i, 11)
lossest_ticker = Cells(i, 9)

End If

If Cells(i, 12) > totalest_value Then
totalest_value = Cells(i, 12)
totalest_ticker = Cells(i, 9)

End If

Next i

Cells(2, 16) = gainest_ticker
Cells(2, 17) = gainest_value
Cells(3, 16) = lossest_ticker
Cells(3, 17) = lossest_value
Cells(4, 16) = totalest_ticker
Cells(4, 17) = totalest_value

Next j




End Sub


