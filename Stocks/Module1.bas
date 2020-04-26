Attribute VB_Name = "Module1"

Sub Stocks()
'Loop through all worksheets
Dim ws As Worksheet
    
For Each ws In Worksheets

    ws.Activate

'Create Column Headings
    Dim Headings(4) As String

    Headings(0) = "Ticker"
    Headings(1) = "Yearly Change"
    Headings(2) = "Yearly Percent Change"
    Headings(3) = "Total Stock Volume"

    Range("I1:L1").Value = Headings

    'Assign variables for analysis
    Dim Ticker As String

    Dim i As Long

    Dim openingPrice, closingPrice, yearlyChange, percentYearlyChange As Double

    Dim totalVolume As Double

    j = 2

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    'Starting value of first stock on opening day
    openingPrice = Cells(2, 3).Value

    totalVolume = 0

        For i = 2 To lastRow

        'Sum values of each stock price
            totalVolume = Cells(i, 7) + totalVolume

        'continues code if there is an overflow error because long will not hold values greater than 2,147,483,647
         '   On Error Resume Next

        ' Searches for when the value of the next cell is different than that of the current cell
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                'Get Ticker name and put in table
                Ticker = Cells(i, 1).Value
                
                Cells(j, 9).Value = Ticker
                
            'Find closing price of stock and calculate change in stock price and percent yearly change
                
                closingPrice = Cells(i, 6).Value
                
                'Comfirmation that the correct cell for closing price was assigned
                Cells(i, 6).Interior.ColorIndex = 3
                
                yearlyChange = closingPrice - openingPrice
                
                'Don't try to divide by 0
                If closingPrice > 0 Then
                
                    percentYearlyChange = yearlyChange / closingPrice
                
                End If
                
                Cells(j, 10).Value = yearlyChange
                
                Cells(j, 10).Style = "Currency"
                
                Cells(j, 11).Value = percentYearlyChange
                
                Cells(j, 11).Style = "Percent"
                    
                    'Format color based on positive or negative yearly change
                    If yearlyChange < 0 Then
                    
                        Cells(j, 10).Interior.ColorIndex = 3
                        
                    ElseIf yearlyChange > 0 Then
                    
                        Cells(j, 10).Interior.ColorIndex = 4
                        
                    Else: Cells(j, 10).Interior.ColorIndex = 0
                        
                    End If
            
            'Reassign opening price to next stock ticker
                openingPrice = Cells(i + 1, 3).Value
                
            'Confirmation that the correct cell for opening price was assigned
                Cells(i + 1, 3).Interior.ColorIndex = 4
                
            'Store Volume total of stock in summary table
                Cells(j, 12).Value = totalVolume
                
            'move to next row in summary table
                j = j + 1
                
            'reset volume sum calculator
                totalVolume = 0

            End If

        Next i
Columns("I:L").AutoFit

Next ws

End Sub

