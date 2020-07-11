# VBA-challenge
Week 2 Homework

VBA script:

'Credit Running on multiple sheets: https://www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html
Sub EVERYTHING()
    Dim Sheet As Worksheet
    Application.ScreenUpdating = False
    For Each Sheet In Worksheets
        Sheet.Select
        Call Pleasework
    Next
    Application.ScreenUpdating = True
End Sub

Sub Pleasework()
Dim ticker_name As String

Dim volume_total As Double
volume_total = 0

Dim summary_row As Integer
summary_row = 2

Dim open_price As Double
open_price = Cells(2, 3).value
Dim close_price As Double
Dim yearly_change As Double
yearly_change = 0
Dim percent_change As Double
percent_change = 0


last_row = Cells(Rows.Count, 1).End(xlUp).Row

Range("I1").value = "Ticker"
Range("J1").value = "Yearly Change"
Range("K1").value = "Percent Change"
Range("L1").value = "Total Stock Volume"

Range("O2").value = "Greatest % Increase"
Range("O3").value = "Greatest % Decrease"
Range("O4").value = "Greatest Total Volume"
'Credit Adjusting column width: https://stackoverflow.com/questions/24058774/excel-vba-auto-adjust-column-width-after-pasting-data
Columns("O").AutoFit
Range("P1").value = "Ticker"
Range("Q1").value = "Value"
Columns("Q").AutoFit


For i = 2 To last_row
    If Cells(i + 1, 1).value <> Cells(i, 1).value Then
        ticker_name = Cells(i, 1).value
        volume_total = volume_total + Cells(i, 7).value
        close_price = Cells(i, 6).value
        yearly_change = close_price - open_price
    If open_price = 0 Then
        open_price = 1
    End If
        percent_change = yearly_change / open_price
        Range("I" & summary_row).value = ticker_name
        Range("J" & summary_row).value = yearly_change
        Range("L" & summary_row).value = volume_total
        Range("K" & summary_row).value = percent_change
        open_price = Cells(i + 1, 3).value
        
        summary_row = summary_row + 1
        volume_total = 0
    
    Else
        volume_total = volume_total + Cells(i, 7).value
    
    End If

Next i

For i = 2 To last_row
    If Cells(i, 10) < 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
    Else
        Cells(i, 10).Interior.ColorIndex = 4

    End If
    
Next i
    
'Credit converting to percentage: https://stackoverflow.com/questions/42844778/vba-for-each-cell-in-range-format-as-percentage
For i = 2 To last_row
    If Cells(i, 11) < 100000 Then
        Cells(i, 11).NumberFormat = "0.00%"
    End If
    
Next i

Max = 0
Min = 0

For i = 2 To last_row
    If Cells(i, 11) >= Max Then
        Max = Cells(i, 11).value
        Range("Q2").value = Max
'Credit converting to percentage: https://stackoverflow.com/questions/42844778/vba-for-each-cell-in-range-format-as-percentage
        Range("Q2").NumberFormat = "0.00%"
        Range("P2").value = Cells(i, 9)
    
    ElseIf Cells(i, 11) <= Min Then
        Min = Cells(i, 11).value
        Range("Q3").value = Min
        Range("Q3").NumberFormat = "0.00%"
        Range("P3").value = Cells(i, 9)
    
    End If

Next i

max1 = 0
For i = 2 To last_row
    If Cells(i, 12) >= max1 Then
        max1 = Cells(i, 12).value
        Range("Q4").value = max1
        Range("P4").value = Cells(i, 9)
    
    End If
    
Next i



End Sub


