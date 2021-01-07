Attribute VB_Name = "Module1"
Sub StockMarket()
    'Initalize
    Dim ws As Worksheet
    Set ws = ActiveSheet

    For Each ws In ThisWorkbook.Worksheets

    'Start
    ws.Activate

    'Variables
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    Dim readRow As Long
    readRow = 2

    Dim writeRow As Long
    writeRow = 2

    Dim tickerStart As Boolean
    tickerStart = True

    Dim startPrice As Double
    startPrice = 0

    Dim closePrice As Double
    closePrice = 0

    Dim yearlyChange As Double
    yearlyChange = 0

    Dim percentChange As Double
    percentChange = 0

    Dim totalVolume As Double
    
    totalVolume = 0
    
    'New table variables
    
    Dim greatestIncrease As Double
    greatestIncrease = 0
    
     Dim greatestDecrease As Double
    greatestDecrease = 0
    
     Dim greatestVolume As Double
    greatestVolume = 0
    
    Dim tickerValue As String
    
    

    'Titles
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("O1") = "Ticker"
    ws.Range("P1") = "Value"
    ws.Range("N2") = "Greatest % Increase"
    ws.Range("N3") = "Greatest % Decrease"
    ws.Range("N4") = "Greatest Total Volume"

    'Format
    Columns(10).NumberFormat = "0.00"
    Columns(11).NumberFormat = "0.00%"

    'Read All WS Rows
    For readRow = 2 To lastRow
        'If: Next Ticker = Current Ticker
        If Cells(readRow + 1, "A").Value = Cells(readRow, "A").Value Then

            'Calculate Volume Column
            totalVolume = totalVolume + Cells(readRow, "G").Value

            'Get Yearly Open Price
            If tickerStart = True Then
                startPrice = Cells(readRow, "C")
                tickerStart = False
            End If

        'Else: New Ticker Next Row
        Else

            'Calcualte Last Ticker Row
            closePrice = Cells(readRow, "F")
            yearlyChange = closePrice - startPrice

            If startPrice = 0 Then
                percentChange = 0
            Else
                percentChange = ((closePrice - startPrice) / startPrice)
            End If
            
            If percentChange >= 0 Then
                ws.Cells(writeRow, "J").Interior.ColorIndex = 4
            Else
                ws.Cells(writeRow, "J").Interior.ColorIndex = 3
            
            End If
            
            totalVolume = totalVolume + Cells(readRow, "G").Value
            
            
            'Calculate New Table
            
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                
                Cells(2, "O") = Cells(readRow, "A")
                Cells(2, "P") = greatestIncrease
                
                Cells(2, "P").NumberFormat = "0.00%"
                
                'Only Format P
                
                
            
            End If
            
            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                
                Cells(3, "O") = Cells(readRow, "A")
                Cells(3, "P") = greatestDecrease
                
                Cells(3, "P").NumberFormat = "0.00%"
                
            End If
            
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                
                 Cells(4, "O") = Cells(readRow, "A")
                Cells(4, "P") = greatestVolume
                 
            End If
            
            'Write
            Cells(writeRow, "I") = Cells(readRow, "A")
            Cells(writeRow, "L") = totalVolume
            Cells(writeRow, "J") = yearlyChange
            Cells(writeRow, "K") = percentChange

            'Iterate Write Row
            writeRow = writeRow + 1

            'Reset Variables For Next Ticker
            tickerStart = True
            totalVolume = 0

        End If
        'Iterate Read Row
        Next readRow
    'Iterate Sheet
    Next
    
End Sub
