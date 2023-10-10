Sub market()

    Dim ticker As String
    Dim row As Integer
    Dim close_value As Double
    Dim start_value As Double
    Dim volume_value As Double
    Dim start_date As Double
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    Length = Cells(Rows.Count, 1).End(xlUp).row
    row = 2
    start_date = "20200102"
    
    
    
        For i = 2 To Length
            If Cells(i, 2).Value = start_date And Cells(i, 1).Value = Cells(i + 1, 1).Value Then
                start_value = Cells(i, 3).Value
                volume_value = volume_value + Cells(i, 7).Value
            ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                ticker = Cells(i, 1).Value
                Range("I" & row).Value = ticker
                close_value = Cells(i, 6).Value
                year_change = close_value - start_value
                Range("J" & row).Value = year_change
                percent_change = FormatPercent(((close_value - start_value) / Abs(start_value)))
                Range("K" & row).Value = percent_change
                volume_value = volume_value + Cells(i, 7).Value
                Range("L" & row).Value = volume_value
                volume_value = 0
                row = row + 1
             Else
                volume_value = volume_value + Cells(i, 7).Value
                

            End If
        Next i
     Length2 = Cells(Rows.Count, 9).End(xlUp).row
        For j = 2 To Length2
            If Cells(j, 10).Value > 0 Then
                Cells(j, 10).Interior.ColorIndex = 4
            Else
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
        
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    

    greatest_increase = 0
    greatest_decrease = 0
    greatest_volume = 0
        For i = 2 To Length2
            If Cells(i, 11) > greatest_increase Then
                greatest_increase = Cells(i, 11)
            End If
            If Cells(i, 11) < greatest_decrease Then
                greatest_decrease = Cells(i, 11)
            End If
            If Cells(i, 12) > greatest_volume Then
                greatest_volume = Cells(i, 12)
            End If
        Next i
        
    Cells(2, 15).Value = FormatPercent(greatest_increase)
    Cells(3, 15).Value = FormatPercent(greatest_decrease)
    Cells(4, 15).Value = greatest_volume
        

End Sub
