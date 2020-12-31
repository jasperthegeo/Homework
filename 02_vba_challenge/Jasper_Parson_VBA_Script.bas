Attribute VB_Name = "Module1"
Sub StockAnalysis()
    
    Dim Tickername  As String
    Dim x           As Integer
    Dim Stockvolume As Double
    Dim StockOpening As Double
    Dim StockClosing As Double
    Dim MaxIncrease As Double
    Dim MaxIncreaseTicker As String
    Dim MaxDecrease As Double
    Dim MaxDecreaseTicker As String
    Dim MaxStockVol As Double
    Dim MaxStockVolTicker As String
    
    For Each ws In ActiveWorkbook.Worksheets
        
        lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        x = 2        'Index for Current result output
        
        ws.Cells(1, 9).Value = "Ticker Symbol"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Columns("K").NumberFormat = "0.00%"
        
        ws.Cells(1, 12).Value = "Stock Volume"
        
        Stockvolume = 0
        Tickername = ws.Cells(2, 1).Value        'Initial Ticker name so that 1st iteration of the IF statement is true.
        StockOpening = ws.Cells(2, 3).Value        'Initial value of Stock Opening
        
        For i = 2 To lastrow
            
            If Tickername = ws.Cells(i, 1).Value Then
                Stockvolume = Stockvolume + ws.Cells(i, 7)
                
            Else
                'Calculating Results
                StockClosing = ws.Cells(i - 1, 6)
                
                'Results output
                ws.Cells(x, 9).Value = Tickername
                ws.Cells(x, 10).Value = (StockClosing - StockOpening)
                
                If StockOpening <> 0 Then        'Some Stocks open at Zero
                ws.Cells(x, 11).Value = (StockClosing - StockOpening) / StockOpening
            End If
            
            ws.Cells(x, 12).Value = Stockvolume
            
            'Re-initilisation of variables
            Tickername = ws.Cells(i, 1)
            StockOpening = ws.Cells(i, 3)
            Stockvolume = 0
            x = x + 1
            
        End If
        
    Next i
    
    ws.Columns("K").FormatConditions.Delete        'Deletes any old formatting that may be lingering
    'First Rule, Red if negative yearly closing
    ws.Columns("K").FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
    Formula1:="=0"
    ws.Columns("K").FormatConditions(1).Interior.Color = vbRed
    'Second Rule, Green if positive yearly closing
    ws.Columns("K").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
    Formula1:="=0"
    ws.Columns("K").FormatConditions(2).Interior.Color = vbGreen
    
    ws.Range("K1").FormatConditions.Delete        'Deletes the odd formatting in the title
    
    ws.Range("O1:O2").NumberFormat = "0.00%"        'Formatting for Percentage
    
    SheetMaxIncrease = WorksheetFunction.Max(ws.Columns("K"))
    If SheetMaxIncrease > MaxIncrease Then
        MaxIncrease = SheetMaxIncrease
        MaxIncreaseTicker = Application.WorksheetFunction.XLookup(SheetMaxIncrease, ws.Columns("K"), ws.Columns("I"), False)        'this function only works in Office 365, early versions including Office 2019 do not support this
    End If
    
    SheetMaxDecrease = WorksheetFunction.Min(ws.Columns("K"))
    If SheetMaxDecrease < MaxDecrease Then
        MaxDecrease = SheetMaxDecrease
        MaxDecreaseTicker = Application.WorksheetFunction.XLookup(MaxDecrease, ws.Columns("K"), ws.Columns("I"), False)        'this function only works in Office 365, early versions including Office 2019 do not support this
    End If
    
    SheetMaxStockVol = WorksheetFunction.Max(ws.Columns("L"))
    If SheetMaxStockVol > MaxStockVol Then
        MaxStockVol = SheetMaxStockVol
        MaxStockVolTicker = Application.WorksheetFunction.XLookup(MaxStockVol, ws.Columns("L"), ws.Columns("I"), False)
    End If
    
Next ws

ActiveWorkbook.Worksheets(1).Cells(2, 15).Value = "Greatest % Increase"
ActiveWorkbook.Worksheets(1).Cells(3, 15).Value = "Greatest % Decrease"
ActiveWorkbook.Worksheets(1).Cells(4, 15).Value = "Greatest Total Volume"

ActiveWorkbook.Worksheets(1).Cells(1, 16).Value = "Ticker"
ActiveWorkbook.Worksheets(1).Cells(1, 17).Value = "Value"

ActiveWorkbook.Worksheets(1).Cells(2, 16).Value = MaxIncreaseTicker
ActiveWorkbook.Worksheets(1).Cells(3, 16).Value = MaxDecreaseTicker
ActiveWorkbook.Worksheets(1).Cells(4, 16).Value = MaxStockVolTicker

ActiveWorkbook.Worksheets(1).Cells(2, 17).Value = MaxIncrease
ActiveWorkbook.Worksheets(1).Cells(3, 17).Value = MaxDecrease
ActiveWorkbook.Worksheets(1).Cells(4, 17).Value = MaxStockVol

ActiveWorkbook.Worksheets(1).Cells(2, 17).NumberFormat = "0.00%"
ActiveWorkbook.Worksheets(1).Cells(3, 17).NumberFormat = "0.00%"

End Sub


