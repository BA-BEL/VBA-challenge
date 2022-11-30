Attribute VB_Name = "Module1"
Sub Summary()

Dim ws As Worksheet

For Each ws In Worksheets

    lastrow = ws.Cells(1, 1).End(xlDown).Row
    
    'Counter for summary table row
    summarycount = 2
    
    'Starter counter for each ticker type
    startcount = 2
    
    'Assign counter for end of each ticker type
    endcount = 0
    
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

'Autofit columns
ws.Columns("A:L").AutoFit

Next ws



End Sub

Sub test_code()

'Test code on singular worksheet before applying it in a for loop

Set ws = Sheets(1)

lastrow = ws.Cells(1, 1).End(xlDown).Row

'Counter for summary table row
summarycount = 2

'Starter counter for each ticker type
startcount = 2

'Assign counter for end of each ticker type
endcount = 0

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

'Autofit columns
ws.Columns("A:M").AutoFit

End Sub
