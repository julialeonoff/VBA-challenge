Attribute VB_Name = "Module1"
Sub VBA_HW():

'Apply to every worksheet
Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate

'Find last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Define variables
Dim ticker As String
Dim yearly_change As Double
Dim open_price As Double
    open_price = ws.Cells(2, 3).Value
Dim close_price As Double
Dim percent_change As Double
Dim volume As Double
volume = 0
Dim summary_table_row As Integer
summary_table_row = 2

'Make table headings
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Volume"

'Loop through the sheets
For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    ticker = ws.Cells(i, 1).Value
    
    close_price = ws.Cells(i, 6).Value
    
    yearly_change = close_price - open_price
    
    If open_price = 0 Then
        percent_change = 0
        Else
        percent_change = yearly_change / open_price
        End If
    
    volume = volume + ws.Cells(i, 7).Value
    
    ws.Range("I" & summary_table_row).Value = ticker
    ws.Range("J" & summary_table_row).Value = yearly_change
    ws.Range("K" & summary_table_row).Value = percent_change
    ws.Range("L" & summary_table_row).Value = volume
    
    summary_table_row = summary_table_row + 1
    
    open_price = ws.Cells(i + 1, 3).Value
    
    volume = 0
    
    Else
    
    volume = volume + ws.Cells(i, 7).Value
    
    End If
    
    Next i
    
    'Make percent change column % format
    ws.Columns("K").NumberFormat = "0.00%"
    
    'Find last row for yearly_change
    yearly_change_LastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
    'Set colors for yearly_change
    For j = 2 To yearly_change_LastRow
    
        If (ws.Cells(j, 10) > 0 Or ws.Cells(j, 10) = 0) Then
            ws.Cells(j, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(j, 10) < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3
    
    End If
    
    Next j
    
    'Challenge
    
    'Set new table headers
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    For r = 2 To yearly_change_LastRow
    
        'Find greatest % increase
        If ws.Cells(r, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & yearly_change_LastRow)) Then
            ws.Range("P2").Value = ws.Cells(r, 9).Value
            ws.Range("Q2").Value = ws.Cells(r, 11).Value
            
        End If
        
        'Find greatest % decrease
        If ws.Cells(r, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & yearly_change_LastRow)) Then
            ws.Range("P3").Value = ws.Cells(r, 9).Value
            ws.Range("Q3").Value = ws.Cells(r, 11).Value
        
        End If
        
        'Find greatest total volume
        If ws.Cells(r, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & yearly_change_LastRow)) Then
            ws.Range("P4").Value = ws.Cells(r, 9).Value
            ws.Range("Q4").Value = ws.Cells(r, 12).Value
        
        End If
        
        Next r
    
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    Next ws

End Sub

