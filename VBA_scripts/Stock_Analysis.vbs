Attribute VB_Name = "Stock_Analysis"
Sub Stock_Analysis()

Dim CurrTick As String
Dim NextTick As String
Dim StockTick As String
Dim StockVol As Double
Dim RowNum As Long
Dim Summary_Row As Integer
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim AnnualChange As Double
Dim PercentChange As Double
Dim MaxPercent As Double
Dim MaxPercemtTick As String
Dim MinPercent As Double
Dim MinPercentTick As String
Dim MaxVol As Double
Dim MaxVolTick As String
Dim ws As Worksheet

For Each ws In Worksheets
    ws.Activate
    
    'Count used rows
    RowNum = Range("A1", Range("A1").End(xlDown)).Count
    
    'Initiate summary table variables
    Summary_Row = 2
    
    'Sort data by ticker & date. Found syntax at this website for sort code.
        'https://trumpexcel.com/sort-data-vba/
    Range(Range("a2").End(xlToRight), Range("a2").End(xlDown)).Sort key1:=Columns("A"), _
    Order1:=xlAscending, Header:=xlNo
    
    'Create headers
    Range("J1, P1").Value = "Ticker"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "% Change"
    Range("M1").Value = "Volume"
    Range("O1").Value = "Best/Worst Stocks"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("Q1").Value = "Value"
    
    'Format headers
    Range("J1:Q1").HorizontalAlignment = xlHAlignCenter
    Range("J1:Q1").VerticalAlignment = xlVAlignCenter
    Range("O2:O4").HorizontalAlignment = xlHAlignLeft
    
    'loop through stock data
    For r = 2 To RowNum
        CurrTick = Cells(r, 1).Value
        NextTick = Cells(r + 1, 1).Value
            If NextTick <> CurrTick Then
                'Set Stock Ticker
                StockTick = Cells(r, 1).Value
                
                'Add Stock Volume to total
                StockVol = StockVol + Cells(r, 7).Value
                
                'Print stock Ticker
                Range("J" & Summary_Row).Value = StockTick
                
                'capture stock opening price & stock closing price
                OpenPrice = Cells(r - RowCount, 3)
                ClosePrice = Cells(r, 6)
                
                'Compute annual: change & Percent change
                AnnualChange = ClosePrice - OpenPrice
                    
                    'Error handling of new stocks that haven't been traded
                    If OpenPrice = 0 And ClosePrice = 0 Then
                        PercentChange = 0
                    Else
                        PercentChange = (ClosePrice - OpenPrice) / OpenPrice
                    End If
             
                    'Capture stock greatest (+) percent change
                    If PercentChange > MaxPercent Then
                        MaxPercent = PercentChange
                        MaxPercentTick = StockTick
                    End If
                    
                    'Capture stock greatest (-) percent change
                    If PercentChange < MinPercent Then
                        MinPercent = PercentChange
                        MinPercentTick = StockTick
                    End If
                    
                    'Capture stock with greatest volume
                    If StockVol > MaxVol Then
                        MaxVol = StockVol
                        MaxVolTick = StockTick
                    End If
                
                'Print Yearly Change, Percent change, & Stock Volume to table Columns K:M. Used Site below for $ format:
                    'https://doc-archives.microstrategy.com/producthelp/10.5/ReportDesigner/WebHelp/Lang_1033/Content_
                    '/ReportDesigner/custom_number_formatting_examples.htm
                Range("K" & Summary_Row).Value = AnnualChange
                Range("k" & Summary_Row).NumberFormat = "$* #,##0.00;$* -#,##0.00"
                Range("L" & Summary_Row).Value = PercentChange
                Range("L" & Summary_Row).NumberFormat = "0.00%"
                Range("M" & Summary_Row).Value = StockVol
                
                    'Nested IF for conditional formatting Green/Red/No color
                    If AnnualChange > 0 Then
                        Range("K" & Summary_Row).Interior.ColorIndex = 4
                    
                    ElseIf AnnualChange < 0 Then
                        Range("K" & Summary_Row).Interior.ColorIndex = 3
                    Else
                        Range("K" & Summary_Row).Interior.ColorIndex = 0
                    End If
                
                'iterate row
                Summary_Row = Summary_Row + 1
                
                'set stock volume to 0
                StockVol = 0
                
                'set RowCount to 0
                RowCount = 0
            Else
                'Adding/Totaling stock volumes
                StockVol = StockVol + Cells(r, 7).Value
                
                'Error handling for OpenPrices = 0
                If Cells(r, 3).Value = 0 Then
                    RowCount = RowCount
                Else
                    RowCount = RowCount + 1
                End If
            End If
    Next r
    
    'Populate best & Worst Stock data
    Range("P2").Value = MaxPercentTick
    Range("Q2").Value = MaxPercent
    Range("P3").Value = MinPercentTick
    Range("Q3").Value = MinPercent
    Range("P4").Value = MaxVolTick
    Range("Q4").Value = MaxVol
    Range("Q2:Q3").NumberFormat = "0.00%"
    Range("J:M").Columns.AutoFit
    
    'Autofit columns
    Range("J:M").Columns.AutoFit
    Range("O:Q").Columns.AutoFit
    
    'Reset Best & Worst Stock data
    MaxPercentTick = ""
    MaxPercent = 0
    MinPercentTick = ""
    MinPercent = 0
    MaxVolTick = ""
    MaxVol = 0
    
Next ws

'Set focus on 1st dataset
Sheet1.Activate

End Sub


