Attribute VB_Name = "Module4"
Sub twoYearStock()

Dim ws As Worksheet
    For Each ws In Worksheets
    
        Dim i As Long
        Dim lastrow As Long
        Dim l As Long
        Dim counter As Long
        Dim StockName As String
        Dim ResultTableRow As Long
        Dim QuarterlyChange As Double
        Dim PercentChange As Double
        Dim TotalStockVolume As LongLong
        Dim StockCounter As Long
        Dim StockOpen As Double
        Dim StockClose As Double
        Dim maxValue As Double
        
        
        ResultTableRow = 2
        QuarterlyChange = 0
        TotalStockVolume = CLng(0)
        StockCounter = 0
        StockOpen = 0
        StockClose = 0
        maxValue = 0
        
        
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        l = 2

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest & Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        
        
        k = lastrow
        
        counter = 0
           
        For i = 2 To k
           
            If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
            
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                StockOpen = ws.Cells(i - StockCounter, 3).Value
                StockClose = ws.Cells(i, 6).Value
                StockName = ws.Cells(i, 1)
                QuarterlyChange = StockClose - StockOpen
                PercentChange = (StockClose - StockOpen) / StockOpen
                
                ws.Range("I" & ResultTableRow).Value = StockName
                ws.Range("L" & ResultTableRow).Value = TotalStockVolume
                ws.Range("J" & ResultTableRow).Value = QuarterlyChange
                ws.Range("K" & ResultTableRow).Value = FormatPercent(PercentChange)
                
                
                ResultTableRow = ResultTableRow + 1
                TotalStockVolume = 0
                StockCounter = 0
                
                
            Else
            
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            StockCounter = StockCounter + 1
            
            
            End If
            counter = counter + 1
            
            
            If ws.Cells(i, 10) < 0 Then
            
                ws.Cells(i, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 4
                
                
            End If
        
        
    
        Next i
        
        ws.Range("P2").Value = maxValue
        MsgBox ("Finished processing " & ws.Name) ' Notify when done with the worksheet
    
    Next ws

End Sub

    

