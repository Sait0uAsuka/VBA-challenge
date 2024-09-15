Sub Multipleyearstockdata()

 Dim ws As Worksheet
 Dim i As Long 'row
 Dim j As Long 'column
 Dim k As Long 'for second loop
 
 For Each ws In Worksheets 'this is for loop to make sure one click of run can run entire worksheet
    Dim lastRow As Long  'this is for last row
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
 
    Dim totalcharge As Double
    totalcharge = 0
 
    Dim openprice As Double
    Dim closeprice As Double
    Dim groupstar As Long
 
 
    Dim percentChange As Double
 
 
 'try to find the highest and lowest value
    Dim maxValue As Double
    Dim minValue As Double
    Dim Geatest_totalValue As Double
 

 
 'Title Name
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Quarterly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"
    ws.Cells(1, 18).Value = "Ticker"
    ws.Cells(1, 19).Value = "Value"
    ws.Cells(2, 17).Value = "Greatest % Increase"
    ws.Cells(3, 17).Value = "Greatest % Decrease"
    ws.Cells(4, 17).Value = "Greatest Total Volume"
 
 
    j = 2                 'column start from 2
    groupstar = 2         'Group star from 2 for calculating which open price started

 
    For i = 2 To lastRow  'nest for loop, for row from 2 to the last row
    

    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then  'if A2 is not equal to A3, then ... "(A2) and (A3) is just an example"
     
     'total sum
     totalcharge = totalcharge + ws.Cells(i, 7).Value
     
     'open price - close price
     openprice = ws.Cells(groupstar, 3).Value
     closeprice = ws.Cells(i, 6).Value
    
     'percentchange show up
            If openprice <> 0 Then
                percentChange = ((closeprice - openprice) / openprice)
            Else
                percentChange = 0
            End If
                
     'Under these lines, it just make sure all the new value will be place into a correct row and column with correct format
     ws.Cells(j, 10).Value = ws.Cells(i, 1).Value
     ws.Cells(j, 11).Value = closeprice - openprice
     ws.Cells(j, 11).NumberFormat = "0.00"
     ws.Cells(j, 12).Value = percentChange
     ws.Cells(j, 12).NumberFormat = "0.00%"
     ws.Cells(j, 13).Value = totalcharge
     totalcharge = 0
     
    'for colour changing
            If ws.Cells(j, 11).Value < 0 Then
                ws.Cells(j, 11).Interior.ColorIndex = 3
            ElseIf ws.Cells(j, 11).Value > 0 Then
                ws.Cells(j, 11).Interior.ColorIndex = 4
            Else
                ws.Cells(j, 11).Interior.ColorIndex = 0
            End If
     
    

             
     j = j + 1              'both of the line is just keep the loop going to the next j and groupstart
     groupstar = i + 1
    
    
    Else 'if the have the same ticker name, total charge will keep adding together until its not match
     
     totalcharge = totalcharge + ws.Cells(i, 7).Value
     closeprice = 0
     openprice = 0
    
    End If
    
 Next i
 

    'my logic here is try to find the MAX and MIN from column L and Column M and place to the correct place (column S)
    'after i got the value, i want to use if condition, if the value is equal to the value on column L
    'then print the exactly ticker name on Column R. Because Column S has the same value of column L or M
    'so it is easy to find the correct ticker name
    
    
    ' Last row for column L and M
    Dim lastRowL As Long
    lastRowL = ws.Cells(ws.Rows.Count, 12).End(xlUp).Row ' For column L
    
    Dim lastRowM As Long
    lastRowM = ws.Cells(ws.Rows.Count, 13).End(xlUp).Row ' For column M
    
    
    'try to get he Max and Min from the entire column L or M
    maxValue = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastRowL))
    minValue = Application.WorksheetFunction.Min(ws.Range("L2:L" & lastRowL))
    Geatest_totalValue = Application.WorksheetFunction.Max(ws.Range("M2:M" & lastRowM))

    

    'Column S, presenting the Max & Min Value from Column L and M
    'this is Greatest % Increase
    ws.Cells(2, 19).Value = maxValue
    ws.Cells(2, 19).NumberFormat = "0.00%"

    'this is Greatest % Decrease
    ws.Cells(3, 19).Value = minValue
    ws.Cells(3, 19).NumberFormat = "0.00%"

    'this is Greatest Total Value
    ws.Cells(4, 19).Value = Geatest_totalValue
 

    ' for telling the Correct Ticker corresponding to the Greatest or Lowest Value
    For k = 2 To lastRowL
        If ws.Cells(2, 19).Value = ws.Cells(k, 12).Value Then
            ws.Cells(2, 18).Value = ws.Cells(k, 10).Value
        End If
    
        If ws.Cells(3, 19).Value = ws.Cells(k, 12).Value Then
            ws.Cells(3, 18).Value = ws.Cells(k, 10).Value
        End If
    Next k

    For k = 2 To lastRowM
            If ws.Cells(4, 19).Value = ws.Cells(k, 13).Value Then
            ws.Cells(4, 18).Value = ws.Cells(k, 10).Value
        End If
    Next k
 
    
    'when i run the code. i found out some name got cut.
    'Hence, I googled and found this code to change the column width to make the name and value be able to see fully.
    ws.Columns("M").ColumnWidth = 20
    ws.Columns("Q").ColumnWidth = 21
    ws.Columns("K").ColumnWidth = 15
    ws.Columns("L").ColumnWidth = 15
    ws.Columns("S").ColumnWidth = 20

 Next ws

End Sub

