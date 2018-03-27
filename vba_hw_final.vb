Sub stock_hard():
    
    ' With ActiveSheet
For Each ws In Worksheets

    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ' End With

    Dim stock_name As String
  ' Set an initial variable for holding the total per stock
    
    Dim stock_Total As Double
    stock_Total = 0

  ' Keep track of the location for each stock in the summary table
    Dim Summary_Table_Row As Integer
    Dim Summary_Table_Row_1 As Integer
    Dim Summary_Table_Row_3 As Integer
    Summary_Table_Row = 2
    Summary_Table_Row_1 = 3
    Summary_Table_Row_3 = 2
    
    Dim open_stock As Double
    Dim close_stock As Double
    
    
    Dim perc_inc As Double
    Dim perc_chg As Double
    
    'Calculating  Total Volume

    For i = 2 To LastRow

    ' Check if we are still within the same stock, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the  name
      stock_name = ws.Cells(i, 1).Value

      ' Add to the Total
      stock_Total = stock_Total + ws.Cells(i, 7).Value

      ' Print the name in the Summary Table
      ws.Range("I" & Summary_Table_Row_3).Value = stock_name

      ' Print the  Amount to the Summary Table
      ws.Range("J" & Summary_Table_Row_3).Value = stock_Total

      ' Add one to the summary table row
      Summary_Table_Row_3 = Summary_Table_Row_3 + 1
      
      ' Reset the Brand Total
      stock_Total = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      stock_Total = stock_Total + ws.Cells(i, 7).Value

    End If

  Next i

    ' Calculating Yearly change and Percent change

  For i = 2 To LastRow
  
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            Set MyMergedRange = ws.Range("a" & i)
            LastRow = MyMergedRange.Row + MyMergedRange.Rows.Count - 1
        
            ws.Range("k2") = ws.Range("c2")

    ' calculating open and cloase stock value  for each stock
    
          first = LastRow + 1
          ws.Range("k" & Summary_Table_Row_1).Value = ws.Cells(first, 3).Value
          ws.Range("L" & Summary_Table_Row).Value = ws.Cells(LastRow, 6).Value
          
          ws.Columns("K:L").EntireColumn.Hidden = True
       
        
          open_stock = ws.Range("k" & Summary_Table_Row).Value
          close_stock = ws.Range("L" & Summary_Table_Row).Value
      If open_stock <> 0 Then
          perc_chg = Round(((close_stock - open_stock) / open_stock) * 100, 2)
        Else
        perc_chg = close_stock
        
       End If
        
          ws.Range("m" & Summary_Table_Row).Value = close_stock - open_stock
          ws.Range("n" & Summary_Table_Row).Value = perc_chg
          
          Summary_Table_Row = Summary_Table_Row + 1
          Summary_Table_Row_1 = Summary_Table_Row_1 + 1
        
        End If

  Next i
  
 
  ' Color Formatting and  Calculating the MAX, MIN ,MAX STOCK VOLUME
  
  ' With ActiveSheet
        LastRow_1 = ws.Cells(ws.Rows.Count, "i").End(xlUp).Row
    ' End With
    
    Dim max As Double
    Dim max_ticker As String
    max = 0
    
     Dim min  As Double
     Dim min_ticker As String
    min = 0
    
    Dim Max_vol As Double
    Dim max__vol_ticker As String
    Max_vol = 0
    

    For j = 2 To LastRow_1
    
        If (ws.Range("n" & j).Value > 0) Then
                ws.Range("m" & j).Interior.ColorIndex = 4
            Else
                ws.Range("m" & j).Interior.ColorIndex = 3
        End If
        
        
         If (ws.Range("N" & j).Value > max) Then
        
           max = ws.Range("N" & j).Value
           max_ticker = ws.Range("i" & j).Value
        End If
        
          If (ws.Range("N" & j).Value < min) Then
        
           min = ws.Range("N" & j).Value
           min_ticker = ws.Range("i" & j).Value
        End If
        If (ws.Range("j" & j).Value > Max_vol) Then
            Max_vol = ws.Range("j" & j).Value
           max__vol_ticker = ws.Range("i" & j).Value
        End If
         
           
    Next j


    ws.Range("p2").Value = max_ticker
    ws.Range("p3").Value = min_ticker
    ws.Range("p4").Value = max__vol_ticker
    ws.Range("q2").Value = Str(max) + "%"
    ws.Range("q3").Value = Str(min) + "%"
    ws.Range("q4").Value = Max_vol
        
Next ws


End Sub














