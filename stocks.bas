Sub stocks():
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Chage"
    Range("K1") = "Percent Chage"
    Range("L1") = "Total Stock Volume"

    Index = 2
    Total = 0
    
    For i = 2 To Cells(Rows.Count, "A").End(xlUp).Row
        Total = Total + Cells(i, "G")
        
        If Cells(i - 1, "A") <> Cells(i, "A") Then
            Openning = Cells(i, "C").Value
        End If

        If Cells(i, "A") <> Cells(i + 1, "A") Then

            Cells(Index, "I") = Cells(i, "A")
            Change = Cells(i, "F").Value - Openning
            Cells(Index, "J") = Change

            If Cells(Index, "J") > 0 Then
                Cells(Index, "J").Interior.ColorIndex = 4
            End If
            
            If Cells(Index, "J") < 0 Then
                Cells(Index, "J").Interior.ColorIndex = 3
            End If
            
            If Openning <> 0 Then
                Cells(Index, "K") = Change / Openning
            End If

            Cells(Index, "K").NumberFormat = "0.00%"
            Cells(Index, "L") = Total
            Cells(Index, "L").NumberFormat = "$ 0,000"
            
            Total = 0
            Index = Index + 1
        End If
    Next i
End Sub


