Attribute VB_Name = "Stock_anal"
Sub Stock_anal()

'To make the tables I need the headers as an array, but making arrays manually
'in VBA is super annoying so I made them as strings and split them into arrays

Dim Titles1 As String
Dim Titles2 As String
Dim Titles3 As String

Titles1 = "Ticker-Yearly Change-Percent Change-Total Stock Volume"
Titles2 = "Ticker-Value"
Titles3 = "Greatest % Increase-Greatest % Decrease-Greeatest Total Volume"

Dim table_1_titles() As String
Dim table_2_titles() As String
Dim table_2_row_titles() As String

table_1_titles = Split(Titles1, "-")
table_2_titles = Split(Titles2, "-")
table_2_rows = Split(Titles3, "-")

Dim i As Long
Dim j As Long
Dim row_count As Long


For EACH ws in Worksheets

    'The following function will construct the output tables.

    'FIrst Table
    For i = 0 To 3
        ws.Cells(1, (i + 9)) = table_1_titles(i)
    Next i

    'Second Table
    For i = 0 To 1
        ws.Cells(1, (i + 16)) = table_2_titles(i)
    Next i

    For i = 0 To 2
        ws.Cells((2 + i), 15) = table_2_rows(i)
    Next i

    'Counting rows

    row_count = 0

    Do While ws.Cells(row_count + 2, 1) <> ""
        row_count = row_count + 1
    Loop

    'Adding 5 to row_count to for some overkill protection against off-by-one errors
    'This was mostly to protect the code through the debugging process, and doesn't add
    'significant time to run, so by gall Im leaving it
    row_count = row_count + 5

    'Populating the first table
    Dim current_symbol As String
    Dim first_price As Double
    Dim last_price As Double
    Dim counter As Long
    counter = 2
    ws.Cells(counter, 12) = 0
    current_symbol = ws.Cells(2, 1)
    first_price = ws.Cells(2, 3)
    For i = 2 To row_count
        ws.Cells(counter, 12) = ws.Cells(counter, 12) + ws.Cells(i, 7)
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
            ws.Cells(counter + 1, 12) = 0
            last_price = ws.Cells(i, 6)
            ws.Cells(counter, 9) = current_symbol
            ws.Cells(counter, 10) = last_price - first_price
            'checks if percentage increase would drop a div/0 error, and if so inputs 0
            If first_price = 0 Then
                ws.Cells(counter, 11) = 0
            Else
                ws.Cells(counter, 11) = ws.Cells(counter, 10) / first_price
            End If
            current_symbol = ws.Cells(i + 1, 1)
            first_price = ws.Cells(i + 1, 3)
            last_price = 0
            counter = counter + 1
        End If
    Next i

    'Populate Second Table
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
        If ws.Cells(i, 11) > gainest_value Then
            gainest_value = ws.Cells(i, 11)
            gainest_ticker = ws.Cells(i, 9)
        ElseIf ws.Cells(i, 11) < lossest_value Then
            lossest_value = ws.Cells(i, 11)
            lossest_ticker = ws.Cells(i, 9)
        End If
        If ws.Cells(i, 12) > totalest_value Then
            totalest_value = ws.Cells(i, 12)
            totalest_ticker = ws.Cells(i, 9)
        End If
    Next i

    ws.Cells(2, 16) = gainest_ticker
    ws.Cells(2, 17) = gainest_value
    ws.Cells(3, 16) = lossest_ticker
    ws.Cells(3, 17) = lossest_value
    ws.Cells(4, 16) = totalest_ticker
    ws.Cells(4, 17) = totalest_value

    'The following is to put proper formatting into the generated tables
    For i=2 to row_count
        ws.Cells(i, 10).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        If ws.Cells(i,10) < 0 Then
            ws.Cells(i,10).Interior.ColorIndex = 3
        ElseIf ws.Cells(i,10) > 0 Then
            ws.Cells(i,10).Interior.ColorIndex = 4
    Next i

    For i =2 to row_count
        ws.Cells(i,11).NumberFormat = "0.00%"
    Next i

    For i = 2 to row_count
        ws.Cells(i,12).NumberFormat = "0"
    Next i

    For i = 2 to 3
        ws.Cells(i,17).NumberFormat = "0.00%"
    Next i

    ws.Cells(4,17).NumberFormat = "0"

Next ws
End Sub


