
Sub vbachal()

For each ws in Worksheets

    Dim Stock As String
    Dim Volume As Double
    Dim First As Double
    Stock = 0
    First = 2
    Dim Summary_Row As Integer
    Summary_Row = 2
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Range("K1:K" & LastRow).Style = "Percent"

    'MsgBox(LastRow)



    For i = 2 To LastRow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        Stock = ws.Cells(i, 1).Value

         

        Volume = Volume + ws.Cells(i, 7).Value
        Y_Change = ws.Cells(i, 6).Value - ws.Cells(First, 3).Value
            
            If ws.Cells(First, 3).Value = 0 Then
            P_Change = 0

            Else
            P_Change = (ws.Cells(i, 6).Value / ws.Cells(First, 3).Value) - 1

            End If

        

        ws.Range("I" & Summary_Row).Value = Stock
        ws.Range("J" & Summary_Row).Value = Y_Change
        ws.Range("K" & Summary_Row).Value = P_Change
        ws.Range("L" & Summary_Row).Value = Volume
        Volume = 0
        Summary_Row = Summary_Row + 1
        First = i + 1

        

        Else

            Volume = Volume + ws.Cells(i, 7).Value

        End If
        
    Next i

    Dim max As Double
    Dim min As Double
    Dim Hi_C As String
    Dim Lo_C As String
    max = 0
    min = 0

    For j = 2 To Summary_Row

        If ws.Cells(j, 11).Value > max Then
        max = ws.Cells(j, 11).Value
        Hi_C = ws.Cells(j, 9).Value
        

        ElseIf ws.Cells(j, 11).Value < min Then
        min = ws.Cells(j, 11).Value
        Lo_C = ws.Cells(j, 9).Value

        Else
        
        End If

    Next j

    Dim Yuge As Double
    Dim Vol_H As String
    Yuge = ws.Cells(2, 12).Value
    Vol_H = ws.Cells(2, 9).Value
    For k = 3 To Summary_Row
        If ws.Cells(k, 12).Value > Yuge Then
            Yuge = ws.Cells(k, 12).Value
            Vol_H = ws.Cells(k, 9).Value
        Else

        End If
    Next k



    ws.Range("Q1:Q2").Style = "Percent"
    ws.Range("O1").Value = "Greatest % Increase"
    ws.Range("O2").Value = "Greatest % Decrease"
    ws.Range("O3").Value = "Highest Volume"
    ws.Range("P1").Value = Hi_C
    ws.Range("P2").Value = Lo_C
    ws.Range("P3").Value = Vol_H
    ws.Range("Q1").Value = max
    ws.Range("Q2").Value = min
    ws.Range("Q3").Value = Yuge
    
Next ws

MsgBox("Run Complete")


End Sub


