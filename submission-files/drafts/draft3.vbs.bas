Attribute VB_Name = "Module3"
Sub Summary_Final2()

Dim ws As Worksheet

For Each ws In Worksheets

        lastrow = ws.Cells(1, 1).End(xlDown).Row
        
        'Counter for summary table row
        summarycount = 2
        
        'Declare variables
        Dim tickername As String
        Dim openprince As Double
        Dim closeprice As Double
        
        '(Total stock volume)
        Dim totalvolume As Double
        
        'Name columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'assign initial open price and total stock volume
        openprice = ws.Cells(2, 3).Value
        totalvolume = 0
        
        For i = 2 To lastrow
        
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
            'Lists ticker symbol
            tickername = ws.Cells(i, 1).Value
            ws.Range("I" & summarycount).Value = tickername
            
            'assign close price
            closeprice = ws.Cells(i, 6).Value
            
            'lists yearly change
            ws.Range("J" & summarycount).Value = closeprice - openprice
            
            'Conditionally formats yearly change (red< 0 & green > 0)
            If ws.Range("J" & summarycount).Value < 0 Then
                ws.Range("J" & summarycount).Interior.ColorIndex = 3
                
                    ElseIf ws.Range("J" & summarycount).Value > 0 Then
                        ws.Range("J" & summarycount).Interior.ColorIndex = 4
                        
                End If
            
            
            'Lists and formats Percent change
            ws.Range("K" & summarycount).Value = (closeprice - openprice) / openprice
            'ws.Range("K" & summarycount).Style = "Percentage" (did not work for some reason; also tried "Percent" and in lower case)
            ws.Range("K" & summarycount).NumberFormat = "0.00%"
            
            'Add stock volume to total
            totalvolume = totalvolume + ws.Range("G" & i).Value
            'List total stock volume
            ws.Range("L" & summarycount).Value = totalvolume
            
            'Reset open price and total stock volume
            openprice = ws.Cells(i + 1, 3).Value
            totalvolume = 0
            
            'Add summarycount row for next entry in the summary
            summarycount = summarycount + 1
            
            Else
            
            'Add stock volume to total
            totalvolume = totalvolume + ws.Range("G" & i).Value
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
    
    summarylastrow = ws.Range("I1").End(xlDown).Row
    
    'Declare new variables (
    Dim tickerin As String
    Dim greatestincrease As Double
    Dim tickerde As String
    Dim greatestdecrease As Double
    Dim tickervol As String
    Dim greatestvolume As Double
    
    'assign initial variables
    greatestincrease = ws.Range("K2").Value
    tickerin = ws.Range("I2").Value
    greatestdecrease = ws.Range("K2").Value
    tickerde = ws.Range("I2").Value
    greatestvolume = ws.Range("L2").Value
    tickervol = ws.Range("I2").Value
    
    For i = 2 To summarylastrow
    
    'check for greatest % increase and print
    If ws.Range("K" & i).Value > greatestincrease Then
        greatestincrease = ws.Range("K" & i).Value
        tickerin = ws.Range("I" & i).Value
        End If
    
    'check for greatest % decrease
    If ws.Range("K" & i).Value < greatestdecrease Then
        greatestdecrease = ws.Range("K" & i).Value
        tickerde = ws.Range("I" & i).Value

        End If
    
    'Check for greatest total volume
    If ws.Range("L" & i).Value > greatestvolume Then
        greatestvolume = ws.Range("L" & i).Value
        tickervol = ws.Range("I" & i).Value
        End If
    
    
    Next i
    
    'Print greatest % increase ticker and value
    ws.Range("P2").Value = tickerin
    ws.Range("Q2").Value = greatestincrease
    
    'Priint greatest % decrease ticker and value
    ws.Range("P3").Value = tickerde
    ws.Range("Q3").Value = greatestdecrease
    
    'Print greatest total volume ticker and value
    ws.Range("P4").Value = tickervol
    ws.Range("Q4").Value = greatestvolume
    
    '----------
    
    'Autofit columns
    ws.Columns("A:R").AutoFit


Next ws


End Sub

