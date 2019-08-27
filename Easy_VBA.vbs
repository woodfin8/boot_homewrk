Sub moderate()

    Dim Stock As String
    Dim Volume As Double
    Dim First As Double
    Stock = 0 
    First = 2
    Dim Summary_Row As Integer
    Summary_Row = 2
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("K1:K"& LastRow).Style = "Percent"

    'MsgBox(LastRow)



    For i = 2 to LastRow

        If Cells(i+1, 1).Value <> Cells(i, 1).Value Then

        Stock = Cells(i,1).Value

         

        Volume = Volume + Cells(i,7).Value
        Y_Change = Cells(i,6).Value - Cells(First,3).Value
            
            If Cells(First,3).Value = 0 Then
            P_Change = 0 

            Else
            P_Change = (Cells(i,6).Value / Cells(First,3).Value) - 1 

            End If 

        

        Range("I"& Summary_Row).Value = Stock
        Range("J"& Summary_Row).Value = Y_Change
        Range("K"& Summary_Row).Value = P_Change
        Range("L"& Summary_Row).Value = Volume
        Volume = 0 
        Summary_Row = Summary_Row + 1

        

        Else

            Volume = Volume + Cells(i,7).Value

        End If


    Next i 

End Sub 