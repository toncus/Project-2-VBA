Sub StockMarketAnalyzer()

Dim long_summary_row As Long
Dim lRow As Long
Dim lCol As Long
Dim long_col As Long
Dim long_row As Long
Dim YrOpen As Single
Dim YrClose As Single
Dim dbl_volume As Double
Dim yropen_first As Long
Dim yropen_last As Long
Dim r_long
Dim Ticker_HighPercent As String
Dim Value_HighPercent As Single
Dim Ticker_LowPercent As String
Dim Value_LowPercent As Single
Dim Ticker_HighVol As String
Dim Value_HighVol As Double
Dim llongRow As Long


Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
    'Last row
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    'Last Column
    lCol = Cells(1, Columns.Count).End(xlToLeft).Column
'Activate new sheet
ws.Activate
Range("I1").Value = "Ticker"
Range("J1").Value = "Total Change"
Range("K1").Value = "               % Change"
Range("L1").Value = "          Volume"
Range("P1").Value = "     Ticker"
Range("Q1").Value = "     Value"
Columns("I:J").ColumnWidth = 15
Columns("K:L").ColumnWidth = 20
Columns("M:N").ColumnWidth = 5
Columns("O:O").ColumnWidth = 20
Columns("P:P").ColumnWidth = 10
Columns("Q:Q").ColumnWidth = 15
Columns("J:J").NumberFormat = "0.00"
Columns("K:K").NumberFormat = "0.00%"
Range("Q2:Q3").NumberFormat = "0.00%"
Range("I1:L1").HorizontalAlignment = xlCenter
long_summary_rows = 2
For long_row = 2 To lRow
    'When on first row
    If Cells(long_row, 1) <> Cells(long_row - 1, 1) Then
            'get year's open
            YrOpen = Cells(long_row, 3)
            yropen_first = long_row
    'When on last row
    ElseIf Cells(long_row, 1) <> Cells(long_row + 1, 1) Then
            'get year's close
           YrClose = Cells(long_row, 6)
           yropen_last = long_row
           'get Ticker
           Cells(long_summary_rows, 9) = Cells(long_row, 1)
           


                'Year Open and YrClose % Change considerations:
                'YrOpen and YrClose are both zero then %age Change equals zero
                If YrOpen = 0 And YrClose = 0 Then
                        Cells(long_summary_rows, 11) = 0
                End If
                'YrOpen is zero but YrClose is something YearOpen defined by 1st date stock traded
                If YrOpen = 0 And YrClose <> 0 Then
                For r_long = yropen_first To yropen_last
                
                    If Cells(r_long, 3) > 0 Then
                        YrOpen = Cells(r_long, 3)
                        Exit For
                    End If
                Next r_long
                
                End If
                

                'If YrOpen is not zero and YrClose is not zero calculate %Change
                If YrOpen <> 0 And YrClose <> 0 Then
                        Cells(long_summary_rows, 11) = ((YrClose - YrOpen) / (YrOpen))
                End If
            'get Year's Total Change
            Cells(long_summary_rows, 10) = YrClose - YrOpen
                'Color Formatting
                If Cells(long_summary_rows, 10) < 0 Then
                        Cells(long_summary_rows, 10).Interior.Color = vbRed
                ElseIf Cells(long_summary_rows, 10) > 0 Then
                        Cells(long_summary_rows, 10).Interior.Color = vbGreen
                Else:
                        Cells(long_summary_rows, 10).Interior.Color = vbWhite
                End If
           'Volume
           Cells(long_summary_rows, 12) = dbl_total + Cells(long_row, 7)

           long_summary_rows = long_summary_rows + 1
           dbl_total = 0
    Else
           'When in middle of ticker rows add to volumes
           dbl_total = dbl_total + Cells(long_row, 7)
    End If

Next long_row
'Summary rows
llongRow = Cells(Rows.Count, 9).End(xlUp).Row
For long_summary_rows = 2 To llongRow

                If long_summary_rows = 2 Then
                    Ticker_HighPercent = Cells(long_summary_rows, 9)
                    Ticker_LowPercent = Cells(long_summary_rows, 9)
                    Ticker_HighVol = Cells(long_summary_rows, 9)
                    Value_HighPercent = Cells(long_summary_rows, 11)
                    Value_LowPercent = Cells(long_summary_rows, 11)
                    Value_HighVol = Cells(long_summary_rows, 12)
                End If
                If Cells(long_summary_rows, 11) > Value_HighPercent Then
                    Value_HighPercent = Cells(long_summary_rows, 11)
                    Cells(2, 15) = "Greatest % Increase"
                    Cells(2, 16) = Cells(long_summary_rows, 9)
                    Cells(2, 17) = Cells(long_summary_rows, 11)
                End If
                If Cells(long_summary_rows, 11) < Value_LowPercent Then
                    Value_LowPercent = Cells(long_summary_rows, 11)
                    Cells(3, 15) = "Greatest % Decrease"
                    Cells(3, 16) = Cells(long_summary_rows, 9)
                    Cells(3, 17) = Cells(long_summary_rows, 11)
                End If
                If Cells(long_summary_rows, 12) > Value_HighVol Then
                    Value_HighVol = Cells(long_summary_rows, 12)
                    Ticker_HighVol = Cells(long_summary_rows, 9)
                
                    Cells(4, 15) = "Greatest Total Volume"
                    Cells(4, 16) = Ticker_HighVol
                    Cells(4, 17) = Value_HighVol
                End If


Next long_summary_rows
'next worsheet
Next ws
'after all worksheets
ActiveWorkbook.Sheets(1).Activate
Range("A1").Select
'Save workbook
ActiveWorkbook.Save
End Sub