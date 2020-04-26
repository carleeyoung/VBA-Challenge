Attribute VB_Name = "Module2"
Sub ChallengeTable():

'loop through worksheets
Dim ws As Worksheet

For Each ws In Worksheets

ws.Activate

'Assign row and column table labels
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'Declare variables
Dim i, lastRow As Integer

Dim max, min, totalVol As Double

Dim tickerMax, tickerMin, tickerVol As String

lastRow = Cells(Rows.Count, 9).End(xlUp).Row

max = Cells(2, 11).Value

min = Cells(2, 11).Value

totalVol = Cells(2, 12).Value

For i = 2 To (lastRow - 1)

    If Cells(i + 1, 11).Value > max Then
        
        max = Cells(i + 1, 11).Value
        
        tickerMax = Cells(i + 1, 9).Value
        
    End If
    
Next i
    
For i = 2 To (lastRow - 1)
    
    If Cells(i + 1, 11).Value < min Then
        
        min = Cells(i + 1, 11).Value
        
        tickerMin = Cells(i + 1, 9).Value
        
    End If
    
Next i
    
For i = 2 To (lastRow - 1)
    
    If Cells(i + 1, 12).Value > totalVol Then
        
        totalVol = Cells(i + 1, 12).Value
        
        tickerVol = Cells(i + 1, 9).Value
        
    End If
    
Next i

'Range("K2:K" & lastRow).Value.max = max
'Range("K2:K" & lastRow).Value.max = min
'Range("L2:L" & lastRow).Value.max = totalVol

Range("P2").Value = tickerMax
Range("P3").Value = tickerMin
Range("P4").Value = tickerVol

Range("Q2").Value = max
Range("Q2").Style = "Percent"
Range("Q3").Value = min
Range("Q3").Style = "Percent"
Range("Q4").Value = totalVol

Columns("O:Q").AutoFit

Next ws

End Sub
