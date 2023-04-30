Sub Ticker()
'This will take care of running all the sheets
    For Each WS In Worksheets
        WS.Activate
 'This will create the headers for all the columns
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        Cells(2, "O").Value = "Greatest % Increase"
        Cells(3, "O").Value = "Greatest % Decrease"
        Cells(4, "O").Value = "Greatest Total Volume"
        Cells(1, "P").Value = "Ticker"
        Cells(1, "Q").Value = "Value"
'Adjust cell cell size
        Range("J1").EntireColumn.AutoFit
        Range("K1").EntireColumn.AutoFit
        Range("O1").EntireColumn.AutoFit
 
'full loop
        Dim i As Long
        Dim MaxPercentChange As Double
        Dim MaxRow As Long
        TotalVolume = 0
        'Points to first row of ticker that has the symbol
        C_Pointer = 2
        Summary_Pointer = 2
        YearlyChange = 0
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        

        
       'is the origninal value plus 1 is assigned  back to label.
        For i = 2 To RowCount
            If Cells(i, "A").Value = Cells(i + 1, "A").Value Then
                TotalVolume = TotalVolume + Cells(i, "G").Value
            Else
                TotalVolume = TotalVolume + Cells(i, "G").Value
                OpenPrice = Cells(C_Pointer, "C").Value
                ClosePrice = Cells(i, "F").Value
                
                YearlyChange = ClosePrice - OpenPrice
                PercentageChange = YearlyChange / OpenPrice * 100
        'Place values in the colmns needed
                Cells(Summary_Pointer, "I").Value = Cells(i, "A").Value
                Cells(Summary_Pointer, "J").Value = YearlyChange
                Cells(Summary_Pointer, "K").Value = "%" & PercentageChange
                Cells(Summary_Pointer, "L").Value = TotalVolume
                
                If YearlyChange > 0 Then
                    Cells(Summary_Pointer, "J").Interior.ColorIndex = 4
                ElseIf YearlyChange < 0 Then
                    Cells(Summary_Pointer, "J").Interior.ColorIndex = 3
                Else
                    Cells(Summary_Pointer, "J").Interior.ColorIndex = 2
                End If
'Check if the current percentage change is higher than the previous maximum
                If percentchange > MaxPercentChange Then
                MaxPercentChange = percentchange
                MaxRow = i
' Write ticker and percentage change associated with the maximum percentage change
                Cells(2, "P").Value = Cells(MaxRow, "A").Value ' write ticker to P2
                Cells(2, "Q").Value = Cells(MaxRow, "K").Value ' write percentage change to Q2
        End If

                
                TotalVolume = 0
                C_Pointer = i + 1
                Summary_Pointer = Summary_Pointer + 1
                
           
            End If
        Next i
        
     Next WS

End Sub
