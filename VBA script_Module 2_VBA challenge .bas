Attribute VB_Name = "Module1"
Sub CalculateQuarterChange():

    For Each ws In Worksheets
    
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim TickerCount As Long
        Dim LastRowA As Long
        Dim PercentChange As Double
        
        
        WorksheetName = ws.Name
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change ($)"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
      
        
        TickerCount = 2
        j = 2
            For i = 2 To LastRowA
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(TickerCount, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(TickerCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    If ws.Cells(TickerCount, 10).Value < 0 Then
                    ws.Cells(TickerCount, 10).Interior.Color = vbRed
                
                    Else

                    ws.Cells(TickerCount, 10).Interior.Color = vbGreen
                
                    End If
                    
                    If ws.Cells(j, 3).Value <> 0 Then
                    PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    ws.Cells(TickerCount, 11).Value = Format(PercentChange, "Percent")
                    
                    Else
                    
                    ws.Cells(TickerCount, 11).Value = 0
                    
                    End If
                    
                ws.Cells(TickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                TickerCount = TickerCount + 1
                
                j = i + 1
                
                End If
            
        Next i
        Next ws
            
End Sub

Sub SumarryTable():
  For Each ws In Worksheets
    
        Dim WorksheetName As String
        Dim LastRowI As Long
        Dim MaxVolValue As Double
        Dim MaxPerValue As Double
        Dim MinPerValue As Double
        Dim i As Long
        
        
        WorksheetName = ws.Name
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest total volume"
        
        MaxVolValue = Application.WorksheetFunction.Max(ws.Range("L1:L" & LastRowI))
        MaxPerValue = Application.WorksheetFunction.Max(ws.Range("K1:K" & LastRowI))
        MinPerValue = Application.WorksheetFunction.Min(ws.Range("K1:K" & LastRowI))
        
        ws.Cells(2, 16).Value = ws.Cells(Application.WorksheetFunction.Match(MaxPerValue, ws.Range("K1:K" & LastRowI), 0), 9).Value
        ws.Cells(2, 17).Value = MaxPerValue
        
        ws.Cells(3, 16).Value = ws.Cells(Application.WorksheetFunction.Match(MinPerValue, ws.Range("K1:K" & LastRowI), 0), 9).Value
        ws.Cells(3, 17).Value = MinPerValue
        
        ws.Cells(4, 16).Value = ws.Cells(Application.WorksheetFunction.Match(MaxVolValue, ws.Range("L1:L" & LastRowI), 0), 9).Value
        ws.Cells(4, 17).Value = MaxVolValue
        
        ws.Cells(2, 17).Value = Format(MaxPerValue, "Percent")
        ws.Cells(3, 17).Value = Format(MinPerValue, "Percent")
        
    Next ws

End Sub
