Sub Stock_Calculations()

'Unit 2 | Assignment - The VBA of Wall Street
'Code by Christopher Reutz

'Initialize variables
Dim DataRowNumber, ResultRowNumber, LastRow As Long
Dim TotalStockVolume, OpenPrice, ClosePrice As Double
Dim GrtVolume, GrtIncrease, GrtDecrease As Double
Dim TckrGrtInc, TckrGrtDec, TckrGrtVol As String
Dim ws As Worksheet

'Loop through each worksheet
For Each ws In Worksheets
    
    'Create header row for the results
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    'CREATE TICKER SUMMMARY RESULTS TABLE - Easy and Moderate Parts
    
    'Initialize variables
    DataRowNumber = 2
    ResultRowNumber = 2
    TotalStockVolume = 0
    
    'Initialize the first open price
    OpenPrice = ws.Cells(DataRowNumber, 3).Value
            
            
    'Find the last row in the data source
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    'Check every row in the data source
    For i = 2 To LastRow

        TotalStockVolume = TotalStockVolume + ws.Cells(DataRowNumber, 7).Value
        
        'Check to see if there ticker will change in the next iteration
        If ws.Cells(DataRowNumber, 1).Value <> ws.Cells(DataRowNumber + 1, 1).Value Then
            
            'Copy ticker value over to the result row
            ws.Cells(ResultRowNumber, 9) = ws.Cells(DataRowNumber, 1).Value
            
            'Set the closing price for the last instance of the stock
            ClosePrice = ws.Cells((DataRowNumber), 6).Value
                      
            'Fill in results for yearly change and color red if negative or green if positive
            ws.Cells(ResultRowNumber, 10).Value = ClosePrice - OpenPrice
            If ws.Cells(ResultRowNumber, 10).Value > 0 Then
                ws.Cells(ResultRowNumber, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(ResultRowNumber, 10).Value < 0 Then
                ws.Cells(ResultRowNumber, 10).Interior.ColorIndex = 3
            End If
                                   
            'Fill in results for the percentage change
            'Set to 0% if open price is 0; avoid a divide by zero error
            If OpenPrice = 0 Then
                ws.Cells(ResultRowNumber, 11).Value = Format(0, "Percent")
            Else
                ws.Cells(ResultRowNumber, 11).Value = Format(((ClosePrice - OpenPrice) / OpenPrice), "Percent")
            End If
            
            'Fill in results the total stock volume
            ws.Cells(ResultRowNumber, 12).Value = TotalStockVolume
            
            'Set the new open price for the next stock
            OpenPrice = ws.Cells(DataRowNumber + 1, 3).Value
            
            'Initialize the stock volume for the next stock
            TotalStockVolume = 0
            
            'Increment new result row
            ResultRowNumber = ResultRowNumber + 1
            
        End If
        
        'Move to the next row in the data
        DataRowNumber = DataRowNumber + 1

    Next i
    
    '--------------------------------------------------------------------
    
    'CREATE GREASTEST INCREASE, DECREASE, VOLUME TABLE - Hard Part
    
    'Initialize variables
    ResultRowNumber = 2
    GrtVolume = 0
    GrtIncrease = 0
    GrtDecrease = 0
    
    'Find the last row in the data source
    LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Check every row in the result set
    For j = 2 To LastRow
        
        'To to see if greatest % increase needs to update
        If ws.Cells(ResultRowNumber, 11).Value > GrtIncrease Then
            GrtIncrease = ws.Cells(ResultRowNumber, 11).Value
            TckrGrtInc = ws.Cells(ResultRowNumber, 9).Value
        End If
        
        'To to see if greatest % decrease needs to update
        If ws.Cells(ResultRowNumber, 11).Value < GrtDecrease Then
            GrtDecrease = ws.Cells(ResultRowNumber, 11).Value
            TckrGrtDec = ws.Cells(ResultRowNumber, 9).Value
        End If
        
        'To to see if greatest total volume needs to update
        If ws.Cells(ResultRowNumber, 12).Value > GrtVolume Then
            GrtVolume = ws.Cells(ResultRowNumber, 12).Value
            TckrGrtVol = ws.Cells(ResultRowNumber, 9).Value
        End If
        
        'Increment new result row
        ResultRowNumber = ResultRowNumber + 1
        
    Next j
        
        'Print greatest results
        ws.Range("P2").Value = TckrGrtInc
        ws.Range("Q2").Value = Format(GrtIncrease, "Percent")
        ws.Range("P3").Value = TckrGrtDec
        ws.Range("Q3").Value = Format(GrtDecrease, "Percent")
        ws.Range("P4").Value = TckrGrtVol
        ws.Range("Q4").Value = GrtVolume
        

'Switch to the next worksheet
Next ws

End Sub
