Attribute VB_Name = "Module1"
Sub stockmarket():

'creating and setting variables
    Dim totalStockVolume As Double
    Dim inforow As Long
    Dim opening As Double
    
        
    totalStockVolume = 0
    inforow = 2
    opening = 0
    
'set column values to hold names

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

'set varibales to capture the entire row

    lastrow = Cells(Rows.Count, 2).End(xlUp).Row
    
'For loop to iterate through entire row

    For rownum = 2 To lastrow
        
        totalStockVolume = totalStockVolume + Cells(rownum, 7).Value
        
        If Cells(rownum, 2).Value = 20160101 Then
            opening = Cells(rownum, 3).Value
      
        End If
    
        If Cells(rownum + 1, 1).Value <> Cells(rownum, 1).Value Then
        
            Cells(inforow, 9).Value = Cells(rownum, 1).Value
            Cells(inforow, 12).Value = totalStockVolume
            
            ' yearly change
                Cells(inforow, 10).Value = (Cells(rownum, 3).Value - opening)
            
            ' Percent Change
                Cells(inforow, 11).Value = (Cells(rownum, 3).Value - opening) / opening
                Cells(inforow, 11).Value = Cells(inforow, 11).Value * 1
                Cells(inforow, 11).NumberFormat = "0.00%"
            
            ' Color Formatting
                If Cells(inforow, 10).Value < 0 Then
                
                    Cells(inforow, 10).Interior.ColorIndex = 3
                    
                Else
                
                    Cells(inforow, 10).Interior.ColorIndex = 4
                    
                End If
                                
            
            totalStockVolume = 0
            inforow = inforow + 1
            

            

        End If
    Next rownum
    

End Sub
