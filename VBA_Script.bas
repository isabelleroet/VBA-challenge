Attribute VB_Name = "Module1"
Sub runtesting()
'Run code on multiple sheets at once
    
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call testing
    Next
    ApplicationScreenUpdating = True

End Sub
Sub runbonus()
'Run code on multiple sheets at once
    
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call bonus
    Next
    ApplicationScreenUpdating = True

End Sub

Sub testing()

    Dim ws As Worksheet
    For Each ws In Worksheets

    Dim lastrow, i, Printer As Long
    Dim StockName, Checker As String
    Dim Volume As Double
    
    Printer = 2
    Volume = 0
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Set column titles
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Open"
    Range("K1").Value = "Close"
    Range("L1").Value = "Change"
    Range("M1").Value = "Percent Change"
    Range("N1").Value = "Total Volume"
    
    'Bonus column titles
    Range("R1").Value = "Ticker"
    Range("S1").Value = "Value"
    Range("Q2").Value = "Greatest % Increase"
    Range("Q3").Value = "Greatest % Decrease"
    Range("Q4").Value = "Greatest Total Volume"
    
        For i = 2 To lastrow
            StockName = Cells(i, 1).Value
                If Cells(i - 1, 1).Value <> StockName Then
                Range("I" & Printer).Value = StockName
                Range("J" & Printer).Value = Cells(i, 3).Value
                Volume = Cells(i, 7)
                
                ElseIf Cells(i + 1, 1) <> StockName Then
                Range("K" & Printer).Value = Cells(i, 6).Value
                Range("N" & Printer).Value = Volume
                Printer = Printer + 1
                Volume = 0
                
                Else: Volume = Volume + Cells(i, 7).Value
                
                End If
                
            Next i
            
            'Calculating the change
            Dim j As Integer
            Dim lastrow1 As Long
            
            lastrow1 = Cells(Rows.Count, 9).End(xlUp).Row
            
                For j = 2 To lastrow1
                    Cells(j, 12) = Cells(j, 10) - Cells(j, 11)
                    
                    If Cells(j, 10) <> 0 Then
                        Cells(j, 13) = Cells(j, 12) / Cells(j, 10)
                        Cells(j, 13).Value = FormatPercent(Cells(j, 13))
                
                    End If
                    
                    'Set to red if less than 0
                    If Cells(j, 13).Value < 0 Then
                        Range("M" & j).Interior.ColorIndex = 3
                    Else
                    'Set to green if not less than 0
                        Range("M" & j).Interior.ColorIndex = 4
                    
                    End If
                    
                    Next j
                    
                    Next ws
        
                                              
End Sub

Sub bonus()

    Dim lastrow As Double
    
    lastrow = Cells(Rows.Count, 13).End(xlUp).Row
    
    Range("S2").Value = FormatPercent(Range("S2"))

    Range("S2").Value = WorksheetFunction.Max(Range("M2:M" & lastrow).Value)
    
    Range("S3").Value = FormatPercent(Range("S3"))
    
    Range("S3").Value = WorksheetFunction.Min(Range("M2:M" & lastrow).Value)
    
    Range("S4").Value = WorksheetFunction.Max(Range("N2:N" & lastrow).Value)
    
    'For Loop for Ticker
        
        For i = 2 To lastrow
             If Cells(i, 13).Value = Range("S2").Value Then
                Range("R2").Value = Cells(i, 9).Value
                End If
             If Cells(i, 13).Value = Range("S3").Value Then
                Range("R3").Value = Cells(i, 9).Value
                End If
             If Cells(i, 14).Value = Range("S4").Value Then
                 Range("R4").Value = Cells(i, 9).Value
                End If
            
            Next i

End Sub
