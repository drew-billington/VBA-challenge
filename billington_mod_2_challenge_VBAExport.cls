VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockAnalysis()


    Dim Ticker As String
    Dim Volume As Double
    Dim Summarytablerow As Integer
    
    Dim BOQValue As Double
    Dim EOQValue As Double
    Dim QuarterlyChange As Double
    Dim PercentChange As Double

    For Each ws In Worksheets
    
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            'MsgBox (LastRow)
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'second summary table titles
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
            Summarytablerow = 2
                            
                For i = 2 To lastrow
                                              
                    
                    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                        
                        'add Ticker symbol to summary table COMPLETE
                        Ticker = ws.Cells(i, 1).Value
                            'print Ticker in Summary Table
                            ws.Cells(Summarytablerow, 9).Value = Ticker
                        
                        'calc volume
                        Volume = Volume + ws.Cells(i, 7).Value
                            'print calculated Volume in Summary table
                            ws.Cells(Summarytablerow, 12).Value = Volume
                                               
                        'calc EOQvalue
                        EOQValue = ws.Cells(i, 6).Value
                            'print EOQvalue for check
                                'ws.Cells(Summarytablerow, 15).Value = EOQValue
                                           
                        'calc BOQvalue
                        BOQValue = ws.Cells(i - 61, 3).Value
                            'print BOQvalue for check
                                'ws.Cells(Summarytablerow, 14).Value = BOQValue
                         
                        'calc Quarterly Change
                        QuarterlyChange = EOQValue - BOQValue
                            'print QChange
                            ws.Cells(Summarytablerow, 10).Value = QuarterlyChange
                                                           
                                'Conditional formatting Quarterly Change
                                If QuarterlyChange > 0.004 Then
                                    ws.Cells(Summarytablerow, 10).Interior.ColorIndex = 4
                                ElseIf QuarterlyChange < -0.004 Then
                                    ws.Cells(Summarytablerow, 10).Interior.ColorIndex = 3
                                Else
                                    ws.Cells(Summarytablerow, 10).Interior.ColorIndex = 0
                                End If
                                                               
                                        
                       'calc Percent Change
                        PercentChange = QuarterlyChange / BOQValue
                            'print Percent Change
                            ws.Cells(Summarytablerow, 11).Value = PercentChange
                        
                        'increase summary table row count = prints next values in next row
                        Summarytablerow = Summarytablerow + 1
                        
                        'return volume count to 0 after printing
                        Volume = 0
                                                     
                    Else
                        
                        'volume counting up to last ticker cell per above code
                        Volume = Volume + ws.Cells(i, 7).Value
                    
                    
                    End If
            
                    
                
            
                Next i
                               
            'format columns to appropriate $/%/,
            ws.Range("J:J").NumberFormat = "$0.00"
            ws.Range("K:K").NumberFormat = "0.00%"
            ws.Range("L:L").NumberFormat = "0,000"
                                
            'autosize columns
            ws.Columns("I:L").AutoFit
                                                   
                                                   
                                                   
    Next ws
    
    'MsgBox ("Complete First Table")
    

    'Second summary table values
    For Each ws In Worksheets
    
        lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
            'MsgBox (lastrow)
        
        
        'print max/min values
        ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & lastrow))
        ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & lastrow))
        ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
        
        'print tickers
        IncreaseTicker = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
        DecreaseTicker = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
        VolumeTicker = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)
            'MsgBox (IncreaseTicker)
        
        ws.Range("P2").Value = ws.Cells(IncreaseTicker + 1, 9)
        ws.Range("P3").Value = ws.Cells(DecreaseTicker + 1, 9)
        ws.Range("P4").Value = ws.Cells(VolumeTicker + 1, 9)
        
        'format cells
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "0,000"
        ws.Columns("O:Q").AutoFit
        
        
        
    Next ws

    'MsgBox ("Complete Second Table")

    MsgBox ("Analysis Complete")

End Sub

