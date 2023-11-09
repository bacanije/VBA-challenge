Attribute VB_Name = "Module1"
Sub StockAnalysis()

        Dim ws As Worksheet
        For Each ws In Worksheets
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        Dim Ticker As String
        Dim Total_Volume As Double
        Total_Volume = 0
        Dim LastRow As Double
        
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Percentage_Change As Double

        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Open_Price = ws.Cells(2, 3).Value
        
        For i = 2 To LastRow
            Total_Volume = Total_Volume + ws.Cells(i, "G")
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value

            ws.Range("I" & Summary_Table_Row).Value = Ticker
            ws.Range("L" & Summary_Table_Row).Value = Total_Volume
            
            Close_Price = ws.Cells(i, 6).Value
            Yearly_Change = Close_Price - Open_Price
            
            If Open_Price <> 0 Then
            Percentage_Change = (Yearly_Change / Open_Price)

            End If
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            ws.Range("K" & Summary_Table_Row).Value = (Percentage_Change)
            ws.Range("K" & Summary_Table_Row).NumberFormat = ("0.00%")
            
            If ws.Range("J" & Summary_Table_Row).Value > 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            ElseIf ws.Range("J" & Summary_Table_Row).Value < 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
            
            
            
            Summary_Table_Row = Summary_Table_Row + 1
            Total_Volume = 0
            Open_Price = ws.Cells(i + 1, 3).Value
            
            
            Else
           
            End If
            
        Next i
    
        ws.Cells(2, 17).Value = WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
        ws.Cells(3, 17).Value = WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
        ws.Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
        
        ws.Range("Q2").NumberFormat = ("0.00%")
        ws.Range("Q3").NumberFormat = ("0.00%")
        
        Greatest_Index = WorksheetFunction.Match(ws.Cells(2, 17).Value, ws.Range("K2:K" & LastRow), 0)
        Greatest_Index = WorksheetFunction.Match(ws.Cells(3, 17).Value, ws.Range("K2:K" & LastRow), 0)
        Greatest_Index = WorksheetFunction.Match(ws.Cells(4, 17).Value, ws.Range("L2:L" & LastRow), 0)
        
    Next ws
    
End Sub

