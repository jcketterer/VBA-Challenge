Attribute VB_Name = "Module1"
Sub stockmarket():

         
'Loop through all worksheets

    For Each ws In Worksheets
    
    'setting variable to hold ticker name and last row
        Dim worksheetname As String
             
        
    'Grab Worksheet
        worksheetname = ws.Name
        
    'Adding ticker to first column header
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        
    'creating and setting variables with worksheet loop
    Dim totalStockVolume As Double
    Dim inforow As Double
    Dim opening As Double
    Dim greatestStockVol As Double
    Dim greatestPercentIncrease As Double
    Dim greatestPercentDecrease As Double
    Dim IncTicker As String
    Dim DecTicker As String
    Dim VolTicker As String
    
    totalStockVolume = 0
    inforow = 2
    opening = 0
    greatestStockVol = 0
    greatestPercentIncrease = 0
    greatestPercentDecrease = 0
       
    'Determine Last row throughout all worksheets
    lastrow = ws.Cells(Rows.Count, 2).End(xlUp).Row

    'set loop
    For rownum = 2 To lastrow
    
    'Determining Stock Value Total and tickers
    totalStockVolume = totalStockVolume + ws.Cells(rownum, 7).Value
        
    'If statement to consolidate tickers and total values
     If ws.Cells(rownum, 2).Value = ws.Cells(2, 2) Then
            opening = ws.Cells(rownum, 3).Value
            
    End If
        
        'If statement to add tickets and add totals
         If ws.Cells(rownum + 1, 1).Value <> ws.Cells(rownum, 1).Value Then
        
            ws.Cells(inforow, 9).Value = ws.Cells(rownum, 1).Value
            ws.Cells(inforow, 12).Value = totalStockVolume
            
        ' If statements for BONUS
          
         If greatestStockVol < totalStockVolume Then
            greatestStockVol = totalStockVolume
            
            VolTicker = ws.Cells(rownum, 1).Value
            
         End If
                     
         If greatestPercentIncrease < ws.Cells(inforow, 11).Value Then
            greatestPercentIncrease = ws.Cells(inforow, 11).Value
            
            IncTicker = ws.Cells(rownum, 1).Value

         End If
        
         If greatestPercentDecrease > ws.Cells(inforow, 11).Value Then
            greatestPercentDecrease = ws.Cells(inforow, 11).Value
            
            DecTicker = ws.Cells(rownum, 1).Value
         End If
                    
        ' yearly change
            ws.Cells(inforow, 10).Value = (ws.Cells(rownum, 3).Value - opening)
                        
        ' Percent Change
            ws.Cells(inforow, 11).Value = (ws.Cells(rownum, 3).Value - opening) / opening
            ws.Cells(inforow, 11).Value = ws.Cells(inforow, 11).Value * 1
            ws.Cells(inforow, 11).NumberFormat = "0.00%"
                       
            
        
         ' Color Formatting
                If ws.Cells(inforow, 10).Value < 0 Then
                
                    ws.Cells(inforow, 10).Interior.ColorIndex = 3
                    
                Else
                
                    ws.Cells(inforow, 10).Interior.ColorIndex = 4
                    
                End If
                
            totalStockVolume = 0
            inforow = inforow + 1
                        
        End If
    
    Next rownum
       
    ws.Cells(4, 16).Value = greatestStockVol
    ws.Cells(4, 15).Value = VolTicker
    
    ws.Cells(3, 16).Value = greatestPercentDecrease
    ws.Cells(3, 15).Value = DecTicker
    ws.Cells(3, 16).NumberFormat = "0.00%"
    
        
    ws.Cells(2, 16).Value = greatestPercentIncrease
    ws.Cells(2, 15).Value = IncTicker
    ws.Cells(2, 16).NumberFormat = "0.00%"
    
       
    Next ws
    

End Sub
