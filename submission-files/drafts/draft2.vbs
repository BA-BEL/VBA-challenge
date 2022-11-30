Attribute VB_Name = "Module2"
Sub Summary_Final()

Dim ws As Worksheet

For Each ws In Worksheets

    lastrow = ws.Cells(1, 1).End(xlDown).Row
    
    'Counter for summary table row
    summarycount = 2
    
    'Starter counter for each ticker type
    startcount = 2
    
    'Assign counter for end of each ticker type
    endcount = 0
    
    'Name columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    For i = 2 To lastrow
    
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        'Lists ticker symbol
        ws.Range("I" & summarycount).Value = ws.Cells(i, 1).Value
        
        'lists yearly change
        ws.Range("J" & summarycount).Value = ws.Cells(startcount + endcount, 6).Value - ws.Cells(startcount, 3).Value
        
        'Conditionally formats yearly change (red< 0 & green > 0)
        If Range("J" & summarycount).Value < 0 Then
            Range("J" & summarycount).Interior.ColorIndex = 3
            
                ElseIf Range("J" & summarycount).Value > 0 Then
                    Range("J" & summarycount).Interior.ColorIndex = 4
                    
            End If
        
        
        'Lists and formats Percent change
        ws.Range("K" & summarycount).Value = (ws.Cells(startcount + endcount, 6).Value - ws.Cells(startcount, 3).Value) / ws.Cells(startcount, 3).Value
        'ws.Range("K" & summarycount).Style = "Percentage" (did not work for some reason; also tried "Percent" and in lower case)
        ws.Range("K" & summarycount).NumberFormat = "0.00%"
        
        'Lists total stock volume
        ws.Range("L" & summarycount).Value = Application.WorksheetFunction.Sum(ws.Range("G" & startcount & ":G" & (startcount + endcount)))
        
        'Increase end count for next startcount assignment
        endcount = endcount + 1
        
        'Add new start count for new ticker
        startcount = startcount + endcount
        
        'Reset end count for next ticker
        endcount = 0
        
        'Add summarycount row for next entry in the summary
        summarycount = summarycount + 1
        
        Else
        
        'increase end count
        endcount = endcount + 1
        End If
    
        
    Next i

'--Bonus section:

'Name columns and rows
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'iterate to make the bonus table

summarylastrow = Cells(1, 1).End(xlDown).Row

For i = 2 To summarylastrow

     

Next i

'Autofit columns
ws.Columns("A:R").AutoFit

Next ws



End Sub
