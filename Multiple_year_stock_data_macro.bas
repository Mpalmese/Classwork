Attribute VB_Name = "Module1"
Sub AlphaTest()

For Each ws In Worksheets

    Dim table_value As Integer
    table_value = 2

    Dim ticker As String

    Dim volume As Double
    volume = 0

    Dim lrow As Long
    lrow1 = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim open_value As Double
    Dim close_value As Double
    Dim year_change As Double
    Dim percent_change As Double
     
    ws.Range("I1").Value = "Ticker"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    
    ws.Columns("K").NumberFormat = "0.00%"
    ws.Range("P2:P3").NumberFormat = "0.00%"
    
        For i = 2 To lrow1
        
            percent_change = 0
            
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                open_value = ws.Cells(i, 3).Value
                End If
    
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                ws.Cells(table_value, 9).Value = ticker
                volume = volume + ws.Cells(i, 7).Value
                ws.Cells(table_value, 12).Value = volume
                
                close_value = ws.Cells(i, 6).Value
                year_change = close_value - open_value
                ws.Cells(table_value, 10).Value = year_change
                
                If open_value <> 0 Then
                    percent_change = year_change / open_value
                End If
                
                ws.Cells(table_value, 11).Value = percent_change
                table_value = table_value + 1
                volume = 0
                
            Else: volume = volume + ws.Cells(i, 7).Value
            End If
            
        Next i
        
Dim lrow2 As Long
lrow2 = Cells(Rows.Count, 9).End(xlUp).Row

ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"

Dim greatest_increase As Double
greatest_increase = 0
ws.Range("P2").Value = greatest_increase

Dim greatest_decrease As Double
greatest_decrease = 0
ws.Range("P3").Value = greatest_decrease

Dim greatest_volume As Double
greatest_volume = 0
ws.Range("P4").Value = greatest_value

    For i = 2 To lrow2
    
        If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
        
        If ws.Cells(i, 11).Value > ws.Range("P2").Value Then
            ws.Range("P2").Value = ws.Cells(i, 11).Value
            ws.Range("O2").Value = ws.Cells(i, 9).Value
        ElseIf ws.Cells(i, 11).Value < ws.Range("P3").Value Then
            ws.Range("P3").Value = ws.Cells(i, 11).Value
            ws.Range("O3").Value = ws.Cells(i, 9).Value
        End If
        
        If ws.Cells(i, 12).Value > ws.Range("P4").Value Then
            ws.Range("P4").Value = ws.Cells(i, 12).Value
            ws.Range("O4").Value = ws.Cells(i, 9).Value
        End If
        
        
    Next i

Next ws
End Sub

