Sub Stock_Calculations()

'Unit 2 | Assignment - The VBA of Wall Street
'Code by Christopher Reutz

'Initialize variables
Dim DataRowNumber, ResultRowNumber As Long
Dim TotalStockVolume, OpenPrice, ClosePrice As Double
Dim GrtVolume, GrtIncrease, GrtDecrease As Double
Dim TickerSymbol, TckrGrtInc, TckrGrtDec, TckrGrtVol As String
Dim WrkSht As Worksheet

'Loop through each worksheet
For Each WrkSht In Worksheets
    
    WrkSht.Select
    
    'Create header row for the results
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    'CREATE TICKER SUMMMARY RESULTS TABLE - Easy and Moderate Parts
    
    'Initialize variables
    DataRowNumber = 2
    ResultRowNumber = 2
    
    'Check every row until the ticker cell becomes empty
    Do While Cells(DataRowNumber, 1).Value <> ""
    
        'Reset the stock volume to 0 when the ticker changes
        TotalStockVolume = 0
        
        'Assign the current ticker when the data source changes to a new ticker
        TickerSymbol = Cells(DataRowNumber, 1).Value
        
        'Set the open price of the stock for the year
        OpenPrice = Cells(DataRowNumber, 3).Value
        
        'Copy over the a new ticker to the result area
        Cells(ResultRowNumber, 9) = TickerSymbol
        
        'Continue to add up the total stock volume while the ticker symbol is the same
        Do While Cells(DataRowNumber, 1).Value = TickerSymbol
            TotalStockVolume = TotalStockVolume + Cells(DataRowNumber, 7).Value
            DataRowNumber = DataRowNumber + 1
        Loop
        
        'Set the closing price for the stock for the year
        ClosePrice = Cells((DataRowNumber - 1), 6).Value
        
        'Fill in results for yearly change and color red if negative or green if positive
        Cells(ResultRowNumber, 10).Value = ClosePrice - OpenPrice
        If Cells(ResultRowNumber, 10).Value > 0 Then
            Cells(ResultRowNumber, 10).Interior.ColorIndex = 4
        ElseIf Cells(ResultRowNumber, 10).Value < 0 Then
            Cells(ResultRowNumber, 10).Interior.ColorIndex = 3
        End If

        'Fill in results for percent change and total stock volume
        'Set to 0% if open price is 0; avoid a divide by zero error
        If OpenPrice = 0 Then
            Cells(ResultRowNumber, 11).Value = Format(0, "Percent")
        Else
            Cells(ResultRowNumber, 11).Value = Format(((ClosePrice - OpenPrice) / OpenPrice), "Percent")
        End If
        Cells(ResultRowNumber, 12).Value = TotalStockVolume
        
        'Increment new result row
        ResultRowNumber = ResultRowNumber + 1
    Loop
    
    'CREATE GREASTEST INCREASE, DECREASE, VOLUME TABLE - Hard Part
    
    'Initialize variables
    ResultRowNumber = 2
    GrtVolume = 0
    GrtIncrease = 0
    GrtDecrease = 0
    
    'Check every row of result table until ticker cell becomes empty
    Do While Cells(ResultRowNumber, 9).Value <> ""
        
        'To to see if greatest % increase needs to update
        If Cells(ResultRowNumber, 11).Value > GrtIncrease Then
            GrtIncrease = Cells(ResultRowNumber, 11).Value
            TckrGrtInc = Cells(ResultRowNumber, 9).Value
        End If
        
        'To to see if greatest % decrease needs to update
        If Cells(ResultRowNumber, 11).Value < GrtDecrease Then
            GrtDecrease = Cells(ResultRowNumber, 11).Value
            TckrGrtDec = Cells(ResultRowNumber, 9).Value
        End If
        
        'To to see if greatest total volume needs to update
        If Cells(ResultRowNumber, 12).Value > GrtVolume Then
            GrtVolume = Cells(ResultRowNumber, 12).Value
            TckrGrtVol = Cells(ResultRowNumber, 9).Value
        End If
        
        'Increment new result row
        ResultRowNumber = ResultRowNumber + 1
        
        'Print greatest results
        Range("P2").Value = TckrGrtInc
        Range("Q2").Value = Format(GrtIncrease, "Percent")
        Range("P3").Value = TckrGrtDec
        Range("Q3").Value = Format(GrtDecrease, "Percent")
        Range("P4").Value = TckrGrtVol
        Range("Q4").Value = GrtVolume
        
    Loop

'Switch to the next worksheet
Next

End Sub
