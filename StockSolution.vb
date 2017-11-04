Sub StockSolution()

Dim ws, ws2 As Worksheet
Set ws = Sheets("Stock_data_2016")
Set ws2 = Sheets.Add(After:=Sheets(Sheets.Count))
Dim ticker As String
Dim total As Double
Dim percentChange As Double
Dim change As Double
Dim dailyChange As Double
Dim days As Integer
Dim j As Integer
LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
total = 0
Start = 2
dailyChange = 0
j = 0

Dim maxValue As Double
maxValue = 0

  For i = 2 To LastRow
'   If ticker changes then print results
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

       ' Stores results in variables
        total = total + ws.Cells(i, 7).Value
        change = (ws.Cells(i, 6) - ws.Cells(Start, 3))
        
        percentChange = Round((change / ws.Cells(Start, 3) * 100), 2)
        
        dailyChange = dailyChange + (ws.Cells(i, 4) - ws.Cells(i, 5))

       ' Average change
        days = (i - Start) + 1
        averageChange = dailyChange / days

       ' start of the next stock ticker
        Start = i + 1

       ' print the results to a seperate worksheet
        ws2.Range("A" & 1) = "Ticker"
        ws2.Range("B" & 1) = "Total Change"
        ws2.Range("C" & 1) = "%OfChange"
        ws2.Range("D" & 1) = "AvgDailyChange"
        ws2.Range("E" & 1) = "Total Volume"
        ws2.Range("A" & 2 + j).Value = ws.Cells(i, 1).Value
        ws2.Range("B" & 2 + j).Value = Round(change, 2)
        ws2.Range("C" & 2 + j).Value = "%" & percentChange
        ws2.Range("D" & 2 + j).Value = averageChange
        ws2.Range("E" & 2 + j).Value = total

       ' colors positives green and negatives red
        Select Case change
            Case Is > 0
               ws2.Range("B" & 2 + j).Interior.ColorIndex = 4
            Case Is < 0
                ws2.Range("B" & 2 + j).Interior.ColorIndex = 3
            Case Else
                ws2.Range("B" & 2 + j).Interior.ColorIndex = 0
        End Select

       ' reset variables for new stock ticker
        total = 0
        change = 0
        j = j + 1
        days = 0
        dailyChange = 0
' If ticker is still the same add results
    Else
        total = total + ws.Cells(i, 7).Value
        change = change + (ws.Cells(i, 6) - ws.Cells(i, 3))

       ' change in high and low
        dailyChange = dailyChange + (ws.Cells(i, 4) - ws.Cells(i, 5))
        



   End If
Next i

'Greatest volume

LastRow = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row

 For i = 2 To LastRow
   If ws2.Cells(i, 5).Value > maxValue Then
     maxValue = ws2.Cells(i, 5).Value
   End If
 
 Next i
 
 
   ws2.Range("H" & 2).Value = maxValue
   ws2.Range("G" & 2).Value = "Greatest volume"
   
   
   
   
   
 'Greatest% increase & decrease
  
Dim maxValue1, minValue As Double
maxValue1 = 0
minValue = 0

 For i = 2 To LastRow
   If ws2.Cells(i, 3).Value > maxValue1 Then
     maxValue1 = ws2.Cells(i, 3).Value
   End If
 
 Next i
 
 For i = 2 To LastRow
   If ws2.Cells(i, 3).Value < minValue Then
     minValue = ws2.Cells(i, 3).Value
   End If
 
 Next i
 
 
   ws2.Range("H" & 5).Value = Format(maxValue1, "Percent")
   ws2.Range("H" & 7).Value = Format(minValue, "Percent")
   ws2.Range("G" & 5).Value = "Greatest % increase"
   ws2.Range("G" & 7).Value = "Greatest % decrease"
   
   
'Greatest Avg. Change

Dim maxValue2 As Double
maxValue2 = 0

 For i = 2 To LastRow
   If ws2.Cells(i, 4).Value > maxValue2 Then
     maxValue2 = ws2.Cells(i, 4).Value
   End If
 
 Next i
 
   ws2.Range("H" & 10).Value = maxValue2
   ws2.Range("G" & 10).Value = "Greatest Avg. Change"
   
   
End Sub
