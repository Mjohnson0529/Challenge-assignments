Attribute VB_Name = "Module1"
Sub ticker()
Dim ticker As String
Dim qtrchg As Double
Dim pctchg As Double
Dim vol As LongLong
Dim stock_info As Integer
Dim start As Long
Dim openprice As Double
Dim closeprice As Double
For Each ws In Worksheets
        ws.Range("I1").Value = "Ticker"
        ws.Range("j1").Value = "Quarterly Change"
        ws.Range("k1").Value = "Percent Change"
        ws.Range("l1").Value = "Total Stock Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % increase"
        ws.Range("N3").Value = "Greatest % decrease"
        ws.Range("N4").Value = "Greatest total volume"
        
        qtrchg = 0
        
        pctchg = 0
        
        vol = 0
       
        start = 2
        previousi = 1
        openprice = ws.Cells(start, 3).Value
        stock_info = 2
        ws.Cells.EntireColumn.AutoFit
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastRow2 = ws.Cells(Rows.Count, 10).End(xlUp).Row
        LastRow3 = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
            For i = 2 To lastrow
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    openprice = ws.Cells(start, 3)
                    previousi = i + 1
                    closeprice = ws.Cells(i, 6).Value
                    ticker = ws.Cells(i, 1).Value
                    qtrchg = (ws.Cells(i, 6).Value - openprice)
                    pctchg = qtrchg / (openprice)
                    vol = vol + ws.Cells(i, 7).Value
                    ws.Range("I" & stock_info).Value = ticker
                    ws.Range("J" & stock_info).Value = qtrchg
                    ws.Range("K" & stock_info).Value = pctchg
                    ws.Range("L" & stock_info).Value = vol
                    ws.Range("K" & stock_info).NumberFormat = "0.00%"
                    stock_info = stock_info + 1
                    vol = 0
                    openprice = Cells(previousi, 3).Value
                    start = previousi
                Else
                    vol = vol + ws.Cells(i, 7).Value
                    previousi = i
                End If
                
                
            Next i
            For i = 2 To LastRow2
                If ws.Cells(i, 10) > 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(i, 10).Value < 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 3
                ElseIf ws.Cells(i, 10).Value = 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 0
                
                
                End If
            Next i
            
            ws.Range("P2").Value = Application.WorksheetFunction.Max(ws.Range("K:K"))
            ws.Range("P3").Value = Application.WorksheetFunction.Min(ws.Range("K:K"))
            ws.Range("P4").Value = Application.WorksheetFunction.Max(ws.Range("L:L"))
            ws.Range("P2:P3").NumberFormat = "0.00%"
            
            For i = 2 To LastRow3
                If ws.Cells(i, 11).Value = ws.Range("P2").Value Then
                    ws.Range("O2").Value = ws.Cells(i, 9).Value
                ElseIf ws.Cells(i, 11).Value = ws.Range("P3").Value Then
                    ws.Range("O3").Value = ws.Cells(i, 9).Value
                ElseIf ws.Cells(i, 12).Value = ws.Range("P4").Value Then
                    ws.Range("O4").Value = ws.Cells(i, 9).Value
                End If
                
            Next i
    Next ws
End Sub
Sub remove()
  For Each ws In Worksheets
        ws.Range("I1:P2000").Value = ""
        ws.Range("J:J").Interior.ColorIndex = 0
    Next ws
End Sub
