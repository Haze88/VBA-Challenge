Attribute VB_Name = "Module1"
Sub StockData()
'to loop through all of the worksheets in an Excel Workbook
    For Each ws In Worksheets

             'variable to hold the worksheet name
             Dim worksheetname As String
    
             'to store the name of the worksheet
             worksheetname = ws.Name
    
            'add ticker to column i
            ws.Range("I1").EntireColumn.Insert
    
            'Added ticker name to desired column
            ws.Range("I1").Value = "Ticker"
            
            'added yearly change to column j
            ws.Range("J1").EntireColumn.Insert
            
            'Added the yearly change name to desired column
             ws.Range("J1").Value = "Yearly Change"
                
            
            'add percent change in column K
            ws.Range("K1").EntireColumn.Insert
            
            'Added percent change name to desired column
             ws.Range("k1").Value = "Percent Change"
            
            'add percent change name to column L
            ws.Range("L1").EntireColumn.Insert
            
            'Added total stock volume name to desired column
            ws.Range("L1").Value = "Total Stock Volume"
    
    
            'summary table for the tickers
            Dim SummaryTable As Integer
            SummaryTable = 2
        
        
            'Variable for ticker
            Dim tickername As String
            
            'variable for yearly chanqge
            Dim yearlychange As Double
            yearlychange = 0
            
            
            'variable for percent change
            Dim percentchange As Double
            percentchange = 0
            
            'variable for total stock calue
            Dim totalstockvolume As Variant
            totalstockvolume = 0
        
    
           'get the last row
           lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
           
             
            'create a variable called open price
            Dim openprice As Double
        
            'create a variable called close price
            Dim closeprice As Double
             
    
            'check through all the tickers
            For Row = 2 To lastRow
        
        'calculate the total stock volume
        totalstockvolume = totalstockvolume + ws.Range("G" & Row).Value
        
        If (Cells(Row, 1).Value <> Cells(Row - 1, 1).Value) Then
            'set open price
             openprice = ws.Range("C" & Row).Value
            
        End If
    
    
        'check to see if we are still on the same ticker name
        'if not do the following
        If (Cells(Row, 1).Value <> Cells(Row + 1, 1).Value) Then
    
             'set close price
             closeprice = ws.Range("F" & Row).Value
            
                'set (reset) the ticker name
                tickername = ws.Range("A" & Row).Value
                
                 'add the ticker name to the desired colum
                     ws.Range("I" & SummaryTable).Value = tickername
                     
                    
                    'calculate the percent change from opening to closing price
                    yearlychange = closeprice - openprice
                     
                         'add yearly change total into column J
                            ws.Range("J" & SummaryTable).Value = yearlychange
                     
                        'highlight positive change in green or negative in red
                         If yearlychange >= 0 Then
                         ws.Range("J" & SummaryTable).Interior.Color = vbGreen
             
                         Else
             
                             ws.Range("J" & SummaryTable).Interior.Color = vbRed
             
                     End If
             
             'reset yearly change for next ticker
             
             
                'calculate the percent change from opening to closing price
                If openprice <> 0 Then
             percentchange = (closeprice - openprice) / openprice
         
            Else
            
            percentchange = closeprice
            
            End If
             
             'add yearly change total into column K
             ws.Range("K" & SummaryTable).Value = percentchange
             
             'reset percent change for next ticker
             
                  
             'calculate the total stock volume
             'add total stock column total into column L
             ws.Range("L" & SummaryTable).Value = totalstockvolume
             
             'reset total stock volume for next ticker
             totalstockvolume = 0
             
             
             'adding next ticker value to summary table
             SummaryTable = SummaryTable + 1
             
        End If
          
          
    Next Row
          
          
  Next ws
    
    
    


End Sub

