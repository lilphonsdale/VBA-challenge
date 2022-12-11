'  * The ticker symbol.
'  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'  * The total stock volume of the stock.

Sub Testing()

'run the script on every worksheet

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets

        'Format Tables
        
        Range("i1").Value = "Ticker"
        Range("j1").Value = "Yearly Change"
        Range("k1").Value = "Percent Change"
        Range("l1").Value = "Total Stock Volume"
        Range("p1").Value = "Ticker"
        Range("q1").Value = "Value"
        Range("o2").Value = "Greatest % Increase"
        Range("o3").Value = "Greatest % Decrease"
        Range("o4").Value = "Greatest Total Volume"
        
        'set variable for ticker
        
        Dim Ticker As String
        
        'set variable for Stock Volume
        
        Dim Volume As LongLong
        Volume = 0
        
        'set variable for Summary Table Row
        
        Dim Table_Row As Integer
        
        'start on row 2
        
        Table_Row = 2
        
        'set variable for year's opening price
        
        Dim OpenPrice As Double
        
        'set variable for year's closing price
        
        Dim ClosePrice As Double
        
        'pull the first stock's opening price
        
        OpenPrice = Cells(2, 3).Value
        
            'make a for loop to analyze the stock data
            
            For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
            
            'use an if statement to identify when the ticker changes
            
                If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                
            'assign that cell value to the ticker variable
                
                Ticker = Cells(i, 1).Value
                
            'add the corresponding volume to the volume variable
                
                Volume = Volume + Cells(i, 7).Value
                
            'assign the corresponding closing price to the closing price variable
                
                ClosePrice = Cells(i, 6).Value
                
            'put the stock ticker into the table
                
                Range("i" & Table_Row).Value = Ticker
                
            'enter the yearly change into the table
               
                Range("j" & Table_Row).Value = ClosePrice - OpenPrice
                 If Range("j" & Table_Row).Value > 0 Then
                    Range("j" & Table_Row).Interior.ColorIndex = 4
                    Else
                    Range("j" & Table_Row).Interior.ColorIndex = 3
                    End If
                
            'enter the yearly percent change into the table with conditional formatting
                
                Range("k" & Table_Row).Value = (ClosePrice - OpenPrice) / OpenPrice
                Range("k" & Table_Row).Value = FormatPercent(Range("k" & Table_Row), 2)
                
            'enter the year's volume into the table
                
                Range("l" & Table_Row).Value = Volume
                
            'move to the next row in the table
                
                Table_Row = Table_Row + 1
                
            'pull the next stock's open price for the year into the OpenPrice variable
                
                OpenPrice = Cells(i + 1, 3).Value
                
            'reset the volume counter
                
                Volume = 0
                
             'what to do when the ticker is not changing from one row to the next
                
                Else
            
            'sum up the volume column
            
                Volume = Volume + Cells(i, 7).Value
            
                End If
            
            Next i
        
             'make a for loop to analyze the summary table
            
            For i = 2 To Cells(Rows.Count, 9).End(xlUp).Row
            
            'make variables to hold the relevant values
            
              Dim CurrentMax As Double
              Dim NewMax As Double
              Dim ChampionTicker As String
              Dim CurrentMin As Double
              Dim NewMin As Double
              Dim LosingTicker As String
              Dim MaxVolume As Double
              Dim MuchVolume As Double
              Dim HeavyHitter As String
              
            
                'to find the best % increase - check if the value is greater than the next cell
                
                If Cells(i, 11).Value > Cells(i + 1, 11) Then
                
                'if the value is greater, save it, if not, keep moving
                NewMax = Cells(i, 11).Value
            
                End If
                
                'compare the saved value against the last champion
                
                If NewMax > CurrentMax Then
        
                'if the value is the new champion, save it, if not keep moving
        
                CurrentMax = NewMax
                
                'save the ticker
                
                ChampionTicker = Cells(i, 9).Value
                
                End If
                
                'to find the worst % decrease - check if the value is less than than the next cell
                
                If Cells(i, 11).Value < Cells(i + 1, 1) Then
                
                'if the value is greater, save it, if not, keep moving
                NewMin = Cells(i, 11).Value
            
                End If
                
                'compare the saved value against the last champion
                
                If NewMin < CurrentMin Then
        
                'if the value is the new champion, save it, if not keep moving
        
                CurrentMin = NewMin
                
                'save the ticker
                
                LosingTicker = Cells(i, 9).Value
                End If
                
                'to find the most volume
                
                If Cells(i, 12).Value > Cells(i + 1, 12) Then
                
                'if the value is greater, save it, if not, keep moving
                MuchVolume = Cells(i, 12).Value
            
                End If
                
                'compare the saved value against the last champion
                
                If MuchVolume > MaxVolume Then
        
                'if the value is the new champion, save it, if not keep moving
        
                MaxVolume = MuchVolume
                
                'save the ticker
                
                HeavyHitter = Cells(i, 9).Value
                
                End If
                
            Next i
            
            'enter the Max Percent Change and its ticker into table
            Cells(2, 17).Value = CurrentMax
            Cells(2, 17).Style = "Percent"
            Cells(2, 16).Value = ChampionTicker
            
            'enter the Min Percent Change and its ticker into table
            Cells(3, 17).Value = CurrentMin
            Cells(3, 17).Style = "Percent"
            Cells(3, 16).Value = LosingTicker
            
             'enter the Min Percent Change and its ticker into table
            Cells(4, 17).Value = MaxVolume
            Cells(4, 16).Value = HeavyHitter

Next ws

End Sub

