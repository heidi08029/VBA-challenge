Attribute VB_Name = "Module1"
Sub assignment2()

 

Dim ws As Worksheet

For Each ws In Worksheets

 

Dim ticker As String

Dim Total_volume As Double

Dim Lastrow As Double

Dim summarytablerow As Integer

Dim Open_price As Double

Dim close_price As Double

Dim yearly_change As Double

Dim percent_change As String

 

Dim max_percent As Double

max_percent = 0

Dim min_percent As Double

min_percent = 0

Dim max_ticker_name As String

Dim min_ticker_name As String

Dim max_volume As Double

max_volume = 0

Dim max_volume_ticker As String

 

Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

summarytablerow = 2

Open_price = ws.Cells(2, 3).Value

 

For i = 2 To Lastrow

 

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        ticker = ws.Cells(i, 1).Value

        Total_volume = ws.Cells(i, 7).Value + Total_volume

        close_price = ws.Cells(i, 6).Value

        yearly_change = close_price - Open_price

        percent_change = (yearly_change / Open_price)

       

        ws.Cells(summarytablerow, 9) = ticker

        ws.Cells(summarytablerow, 12) = Total_volume

        ws.Cells(summarytablerow, 10) = yearly_change

        ws.Cells(summarytablerow, 11) = FormatPercent(percent_change, 2)

       

    If (yearly_change > 0) Then

            ws.Range("J" & summarytablerow).Interior.ColorIndex = 4

    ElseIf (yearly_change <= 0) Then

            ws.Range("J" & summarytablerow).Interior.ColorIndex = 3

    End If

   

        

    If (percent_change > max_percent) Then

        max_percent = percent_change

        max_ticker_name = ticker

    End If

   

    If (percent_change < min_percent) Then

        min_percent = percent_change

        min_ticker_name = ticker

    End If

   

    If (Total_volume > max_volume) Then

        max_volume = Total_volume

        max_volume_ticker = ticker

    End If

   

    summarytablerow = summarytablerow + 1

    Total_volume = 0

    percent_change = 0

    Open_price = ws.Cells(i + 1, 3).Value

       

    Else

        Total_volume = ws.Cells(i, 7).Value + Total_volume

    End If

 

Next i

 

ws.Cells(1, 9).Value = "Ticker"

ws.Cells(1, 10).Value = "Yearly Change"

ws.Cells(1, 11).Value = "Percent Change"

ws.Cells(1, 12).Value = "Total Ticker Volume"

       

 

ws.Cells(2, 15).Value = "Greatest % Increase"

ws.Cells(3, 15).Value = "Greatest % Decrease"

ws.Cells(4, 15).Value = "Greatest Total Volume"

 

ws.Cells(2, 16).Value = max_ticker_name

ws.Cells(2, 17).Value = FormatPercent(max_percent, 2)

 

ws.Cells(3, 16).Value = min_ticker_name

ws.Cells(3, 17).Value = FormatPercent(min_percent, 2)

 

ws.Cells(4, 17).Value = max_volume

ws.Cells(4, 16).Value = max_volume_ticker

 

Next ws

 

End Sub
