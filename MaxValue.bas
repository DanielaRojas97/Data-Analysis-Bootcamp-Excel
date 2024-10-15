Attribute VB_Name = "Module7"
Sub maxValue()
   
         Dim ws As Worksheet
    For Each ws In Worksheets
    
        Dim i As Long
        Dim lastrow As Long
        Dim l As Long
        Dim counter As Long
        Dim maxValue As Double
        Dim ticker1 As String
        Dim ticker2 As String
        Dim ticker3 As String
        Dim maxVolume As LongLong
        Dim minValue As Double

    
        
        
        maxValue = 0
        maxVolume = 0
        minValue = 0
        
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        l = 2

        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest & Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        
        
        k = lastrow
        
        counter = 0
           
        For i = 2 To k
           
            If ws.Cells(i, 11).Value > maxValue Then
                maxValue = ws.Cells(i, 11).Value
                ticker1 = ws.Cells(i, 9).Value
            
            
            End If
            
            If ws.Cells(i, 12).Value > maxVolume Then
                maxVolume = ws.Cells(i, 12).Value
                ticker2 = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 11).Value < minValue Then
                minValue = ws.Cells(i, 11).Value
                ticker3 = ws.Cells(i, 9).Value
                
            End If
            
        
    
        Next i
        
        ws.Range("Q2").Value = FormatPercent(maxValue)
        ws.Range("P2").Value = ticker1
        
        ws.Range("Q3").Value = FormatPercent(minValue)
        ws.Range("P3").Value = ticker3
        
        ws.Range("Q4").Value = maxVolume
        ws.Range("P4").Value = ticker2
        MsgBox ("Finished processing " & ws.Name) ' Notify when done with the worksheet
    
    Next ws

End Sub


