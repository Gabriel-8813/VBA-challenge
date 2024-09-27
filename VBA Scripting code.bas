Attribute VB_Name = "Module1"
Sub Ticker_StockAnalysis()

            ' Assign the worksheet
            Dim ws As Worksheet
            Set ws = ThisWorkbook.Sheets("Q1")
        
            ' Set variable to hold ticker name
            Dim ticker As String
        
            ' Set variable to hold yearly change value
            Dim Quarterly_change As Double
            ' Set variable to hold percentage change value
            Dim percentage_change As Double
            ' Set variable to hold volume value
            Dim volume As Double
            ' Set initial volume value to 0
            volume = 0
        
            ' Track location of each value in summary table
            Dim summary_row As Integer
            summary_row = 2
            ' Find the last row in column A
            Dim last_row As Long
            last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            ' Declare first value for the open price for iteration
            Dim open_price As Double
            open_price = ws.Cells(2, 3).Value
        
            ' Create headers for the summary table
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Quarterly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            
            '  Create headers for the functionalities tables
                    ws.Range("P1").Value = "Ticker"
                    ws.Range("Q1").Value = "Value"
            
            ' Initialize variables for greatest values
            Dim greatest_increase As Double
            Dim greatest_decrease As Double
            Dim greatest_volume As Double
            Dim greatest_increase_stock As String
            Dim greatest_decrease_stock As String
            Dim greatest_volume_stock As String
            
            greatest_increase = 0
            greatest_decrease = 0
            greatest_volume = 0
             
    For Each ws In ThisWorkbook.Worksheets
            ' Reset variables for each worksheet
            volume = 0
            summary_row = 2
            greatest_increase = 0
            greatest_decrease = 0
            greatest_volume = 0
            
            ' Set the last row and open price for the current worksheet
            last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            open_price = ws.Cells(2, 3).Value
    
            ' Clear previous results
            ws.Range("I1:L1").ClearContents
            ws.Range("O1:Q4").ClearContents
    
            ' Set headers
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Quarterly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
        
        
        ' Iterate through all rows and populate summary table
        Dim i As Long
    
        For i = 2 To last_row
            ' Add the volume value of current ticker
            volume = volume + ws.Cells(i, 7).Value
            ' Check if the next row is a different ticker or if it's the last row
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Or i = last_row Then
                ' Set current ticker name
                ticker = ws.Cells(i, 1).Value
                
                ' Set closing price for current ticker
                Dim close_price As Double
                close_price = ws.Cells(i, 6).Value
                ' Find yearly change for current ticker
                Quarterly_change = (close_price - open_price)
                ' Find percentage change for current ticker
                If open_price <> 0 Then
                    percentage_change = (Quarterly_change / open_price)
                Else
                    percentage_change = 0
                End If
                ' Populate summary table with found values
                ws.Range("I" & summary_row).Value = ticker
                ws.Range("J" & summary_row).Value = Quarterly_change
                ws.Range("K" & summary_row).Value = percentage_change
                ws.Range("L" & summary_row).Value = volume
                ' Format percentage change
                ws.Range("K" & summary_row).NumberFormat = "0.00%"
                ws.Range("J" & summary_row).NumberFormat = "0.00"
             
                ' Color the percentage change cell based on its value
                Select Case percentage_change
                    Case Is > 0
                        ws.Range("J" & summary_row).Interior.ColorIndex = 4 ' Green
                    Case Is < 0
                        ws.Range("J" & summary_row).Interior.ColorIndex = 3 ' Red
                    Case Else
                        ws.Range("J" & summary_row).Interior.ColorIndex = 0 ' No Color
                End Select
                
                ' Check for greatest % increase, % decrease, and Total volume
                If percentage_change > greatest_increase Then
                    greatest_increase = percentage_change
                    greatest_increase_stock = ticker
                End If
                
                If percentage_change < greatest_decrease Then
                    greatest_decrease = percentage_change
                    greatest_decrease_stock = ticker
                End If
                
                If volume > greatest_volume Then
                    greatest_volume = volume
                    greatest_volume_stock = ticker
                End If
                
                ' If not at the last row, set new open price for a new ticker
                If i <> last_row Then
                    open_price = ws.Cells(i + 1, 3).Value
                End If
                ' Move to next row in summary table
                summary_row = summary_row + 1
                
                ' Reset volume for new ticker
                volume = 0
                
            End If
            
        Next i
                
                ' Populate greatest values in the summary table
                ws.Range("O2").Value = "Greatest % Increase"
                ws.Range("P2").Value = greatest_increase_stock
                ws.Range("Q2").Value = greatest_increase ' Add this line to show the value
                
                ws.Range("O3").Value = "Greatest % Decrease"
                ws.Range("P3").Value = greatest_decrease_stock
                ws.Range("Q3").Value = greatest_decrease ' Add this line to show the value
                
                ws.Range("O4").Value = "Greatest Total Volume"
                ws.Range("P4").Value = greatest_volume_stock
                ws.Range("Q4").Value = greatest_volume ' Add this line to show the value
                
                ' Format percentage change
                ws.Range("Q2:Q4").NumberFormat = "0.00%"
                
                ' Move to next row in summary table
                summary_row = summary_row + 1
        
    Next ws

End Sub
