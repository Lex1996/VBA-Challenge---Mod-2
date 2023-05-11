Attribute VB_Name = "Module1"
Sub calculations()

    For Each ws In Worksheets
    
        Dim ticker As String
        Dim vol, greatvol, greatper, lessper As Double
        Dim summary As Integer
        Dim lastrow, ticker_row As Long
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % increase"
        ws.Range("N3").Value = "Greatest % decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'PRIMARY CALCULATIONS
        
        summary = 2
        ticker_row = 2
        vol = 0
        
        For i = 2 To lastrow
        
           If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                'Populating the ticker name
                ticker = ws.Cells(i, 1).Value
                ws.Range("I" & summary).Value = ticker
                
                'Calculating the Total Volume
                vol = vol + ws.Cells(i, 7).Value
                ws.Range("L" & summary).Value = vol
                
                'Calculating the Yearly Change
                ws.Range("J" & summary).Value = ws.Cells(i, 6).Value - ws.Cells(ticker_row, 3)
                
                'Calculating the Percentage Change
                ws.Range("K" & summary).Value = ws.Range("J" & summary).Value / ws.Cells(ticker_row, 3)
                
                'Reseting the constants
                ticker_row = i + 1
                summary = summary + 1
                vol = 0
                
            Else
            
                vol = vol + ws.Cells(i, 7).Value
                
            End If
            
        Next i
        
        'SECONDARY CALCULATIONS
        
        greatvol = ws.Cells(2, 12).Value
        greatper = ws.Cells(2, 11).Value
        lessper = ws.Cells(2, 11).Value
        
        For i = 2 To lastrow
        
            'Calculating greatest % increase
            If ws.Cells(i, 11).Value > greatper Then
            
                greatper = ws.Cells(i, 11).Value
                ws.Range("P2").Value = greatper
                ws.Range("O2").Value = ws.Cells(i, 9).Value
                
            Else
            
                greatper = greatper
                
            End If
               
            'Calculating greatest % decrease
            If ws.Cells(i, 11).Value < lessper Then
            
                lessper = ws.Cells(i, 11).Value
                ws.Range("P3").Value = greatper
                ws.Range("O3").Value = ws.Cells(i, 9).Value
                
            Else
            
                lessper = lessper
                
            End If
            
            'Calculating greatest total volume
            If ws.Cells(i, 12).Value > greatvol Then
            
                greatvol = ws.Cells(i, 12).Value
                ws.Range("P4").Value = greatvol
                ws.Range("O4").Value = ws.Cells(i, 9).Value
                
            Else
            
                greatvol = greatvol
                
            End If
            
        Next i
        
        'FORMATTING
        
        'Change formatting of percentages
        ws.Range("K2:K" & lastrow).NumberFormat = "0.00%"
        ws.Range("P2:P3").NumberFormat = "0.00%"
        
        'Change the text formating
        ws.Range("I1:P1").Font.Bold = True
        ws.Range("N2:N4").Font.Bold = True
        ws.Range("I1:P1").Font.Bold = True
        
        'Changing row width
        ws.Range("J1:J" & lastrow).ColumnWidth = 12
        ws.Range("K1:K" & lastrow).ColumnWidth = 14
        ws.Range("L1:L" & lastrow).ColumnWidth = 18
        ws.Range("N1:N" & lastrow).ColumnWidth = 22
        
        'Changing the cell colors according to yearly change
        For i = 2 To lastrow
        
            If ws.Cells(i, 10).Value < 0 Then
            
                ws.Cells(i, 10).Interior.ColorIndex = 3
                
            Else
            
                ws.Cells(i, 10).Interior.ColorIndex = 4
                
            End If
            
        Next i
        
     Next ws
        
End Sub
