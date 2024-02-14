Attribute VB_Name = "Module1"
Sub stock_anaylsis()

For Each ws In Worksheets

'column names
ws.Range("I1:L1").Value = Array("Ticker", "Yearly", "Percentage Change", "Total Stock Volume")
ws.Range("P1:Q1").Value = Array("Ticker", "Value")
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatsest Total Volume"

    'Defining
    Dim Ticker As String
    Dim YC, PC, Opennum, Closenum As Double
    YC = 0
    PC = 0
    Dim LR As Long
    Dim TSV As Double
    TSV = 0
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'Last Row
    LR = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Opennum = ws.Cells(2, 3).Value
    
    
    'First loop for Ticker, Total Volume, YC, and PC
    For i = 2 To LR

        ' when the next row is differnt ticker then add row and add to yellow table
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Finding Opennum and YC
            Closenum = ws.Cells(i, 6).Value
            YC = Closenum - Opennum
            
            'Calculating PC
            PC = (YC / Opennum)
            
            'Define Opennum to find next one
            Opennum = ws.Cells(i + 1, 3).Value
        
            'Define variable locations
            Ticker = ws.Cells(i, 1).Value
            
            'add Total Stock Volume/calculate values
            TSV = TSV + ws.Cells(i, 7).Value
            
            'put ticker & total volume & YC & PC in table
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            ws.Range("L" & Summary_Table_Row).Value = TSV
            ws.Range("J" & Summary_Table_Row).Value = YC
            ws.Range("K" & Summary_Table_Row).Value = PC
            
            'percent formating
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
            'Changing Colors
            If YC > 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            ElseIf YC < 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            Else
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 2
            End If
            
            'Add row in yellow table
            Summary_Table_Row = Summary_Table_Row + 1
            
            'reset total
            TSV = 0
            YC = 0
            
            
        'If same ticker in following row then...
        Else
        
        'Add to Total Stock Volume
        TSV = TSV + ws.Cells(i, 7).Value
        
        End If
        

    Next i

   ' Second Loop for Min%, Max%, & Max Total
   
    For j = 2 To LR
   
        Dim MaxVal, MinVal, TotalMaxVal As Double

        If ws.Cells(j, 11).Value > MaxVal Then
            MaxVal = ws.Cells(j, 11).Value
        End If
    
        If ws.Cells(j, 12).Value > TotalMaxVal Then
            TotalMaxVal = ws.Cells(j, 12).Value
        End If
    
    Next j
    
    ws.Cells(2, 17).Value = MaxVal
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(4, 17).Value = TotalMaxVal
    MinVal = MaxVal
    
    For k = 2 To LR

        If ws.Cells(k, 11).Value < MinVal Then
            MinVal = ws.Cells(k, 11).Value
        End If
    
    Next k
    
    ws.Cells(3, 17).Value = MinVal
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    'Third Loop for Associated Ticker Symbol for Min/Max
    'If PC or TSV Value equals Value in Max/Min Table then the Ticker Value in adjacent
    'Ticker Cell will "print" into Ticker Value in Max/Min Table
    
    Dim TickerMax, TickerMin, TickerTotalMax As String
    
    For m = 2 To LR
    
        If ws.Cells(2, 17).Value = ws.Cells(m, 11).Value Then
            TickerMax = ws.Cells(m, 9).Value
            ws.Cells(2, 16).Value = TickerMax
        End If
    
        If ws.Cells(3, 17).Value = ws.Cells(m, 11).Value Then
            TickerMin = ws.Cells(m, 9).Value
            ws.Cells(3, 16).Value = TickerMin
        End If
    
        If ws.Cells(4, 17).Value = ws.Cells(m, 12).Value Then
            TickerTotalMax = ws.Cells(m, 9).Value
            ws.Cells(4, 16).Value = TickerTotalMax
        End If
    
    Next m

Next
   
End Sub
   
