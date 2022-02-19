Attribute VB_Name = "Module3"
Sub ShareMarket_CJB():

'Variable for cycling through WORKSHEETS.

    Dim ws As Worksheet

'Start processes that are to be applied to ALL worksheets.

    For Each ws In Worksheets

'Variables for counting the number of rows in each year's OVERALL data.

    Dim x As Double
    Dim RowNum As Double

'Variables for counting the number of rows in each year's SUMMARY data.

    Dim y As Integer
    Dim DifferentSharesCount As Integer

'Variables for loops

    Dim i As Double
    Dim n As Double
    Dim k As Double
    Dim q As Double
    Dim r As Double

'Variable for determining the number of DIFFERENT shares traded each year.

    Dim ShareCount As Integer

'Variable for Yearly Price Change for a share.

    Dim YearlyChange As Double

'Variable for determining where to display answers.

    Dim OutputCellCount As Double

'Variable for PERCENTAGE Yearly Price Change for a share.

    Dim PercentageChange As Double

'Variable for Total Stock Volume for a share in a given year.

    Dim TotalShareVolume As Double

'Variable for Referencing Summary Tables Area (for formatting).

    Dim SummaryCells As Range
    Dim AdditionalSummaryCells As Range
    Dim AdditionalSummaryCells2 As Range

'Bonus Section variables.

    Dim LargerIncrease As Double
    Dim GreaterDecrease As Double
    Dim LargestVolume As Double

'Establishinging Column Headers

    ws.Range("H1").Value = "Running Tally         "
    ws.Range("H1").Columns.AutoFit

    ws.Range("I1").Columns.ColumnWidth = 10

    ws.Range("J1").Value = "Ticker  "
    ws.Range("J1").Columns.AutoFit

    ws.Range("K1").Value = "Yearly Change  "
    ws.Range("K1").Columns.AutoFit

    ws.Range("L1").Value = "Percentage Change  "
    ws.Range("L1").Columns.AutoFit

    ws.Range("M1").Value = "Total Stock Volume  "
    ws.Range("M1").Columns.AutoFit

    ws.Range("N1").Value = "Days Traded in Year  "
    ws.Range("N1").Columns.AutoFit

    ws.Range("O1").Columns.ColumnWidth = 6
    
    'Additional Table 1

    ws.Range("P1").Value = "Opening Price (This Year)  "
    ws.Range("P1").Columns.AutoFit

    ws.Range("Q1").Value = "Closing Price (This Year)  "
    ws.Range("Q1").Columns.AutoFit

    ws.Range("R1").Columns.ColumnWidth = 10

    ws.Range("S1").Value = "Diff. Shares Traded This Year "
    ws.Range("S1").Columns.AutoFit

    ws.Range("T1").Value = "Rows in Data Set for Year  "
    ws.Range("T1").Columns.AutoFit

'Bonus Table - For Year

    ws.Range("S5").Value = "BONUS SECTION:"

    ws.Range("T7").Value = "For THIS Year:"
    ws.Range("U7").Value = "Ticker:"
    ws.Range("U7").Columns.AutoFit

    ws.Range("S8").Value = "Greatest % Increase:"
    ws.Range("S9").Value = "Greatest % Decrease:"
    ws.Range("S10").Value = "Greatest Stock Volume:"

    
'Determining the number of rows in overall table for a year's data.
    
    x = ws.Range("A" & Rows.Count).End(xlUp).Row
    
'Determining the number of rows of data minus the column headers.
    RowNum = x - 1
    ws.Range("T2").Value = RowNum
    
'Determining the number of days each share was traded in a year.

    ShareCount = 0
    
    OutputCellCount = 1
    
    For i = 2 To x
       
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            'Choosing a row to write the ticker in.
        
            OutputCellCount = OutputCellCount + 1
    
            'Displaying the Ticker in chosen row.
    
            ws.Cells(OutputCellCount, 10).Value = ws.Cells(i, 1).Value
    
            'Display the number of rows that Share appeared in.
    
            ws.Cells(OutputCellCount, 14).Value = ShareCount + 1
    
            'Display the Share opening price.
    
            ws.Cells(OutputCellCount, 16).Value = ws.Cells(i - ShareCount, 3).Value
    
            'Display the Share closing price.
    
            ws.Cells(OutputCellCount, 17).Value = ws.Cells(i, 6).Value
    
            'Reset the ShareCount for next Share.
    
            ShareCount = 0
    
        Else
    
            ShareCount = ShareCount + 1
    
        End If
    
    Next i

'Determining the number of different shares being traded in overall table for a year's data.
    
    y = ws.Range("J" & Rows.Count).End(xlUp).Row
    
    DifferentSharesCount = y - 1
    
    ws.Range("S2").Value = DifferentSharesCount
    
'Determining the Yearly Price Change

    For n = 2 To DifferentSharesCount + 1

    YearlyChange = ws.Cells(n, 17).Value - ws.Cells(n, 16).Value

    ws.Cells(n, 11).Value = YearlyChange

    Next n

'Determining the Percentage Yearly Price Change

    For n = 2 To DifferentSharesCount + 1

    PercentageChange = 100 * ws.Cells(n, 11).Value / ws.Cells(n, 16).Value
    ws.Cells(n, 12).Value = PercentageChange

    Next n

'Creating a Running Tally for the Total Share Volume
        
    OutputCellCount = 2
    TotalShareVolume = ws.Cells(2, 7).Value

    For k = 2 To x - 1
        
    OutputCellCount = OutputCellCount + 1
    
        If ws.Cells(k + 1, 1).Value = ws.Cells(k, 1).Value Then
    
            TotalShareVolume = TotalShareVolume + ws.Cells(k + 1, 7).Value
            ws.Cells(OutputCellCount, 8).Value = TotalShareVolume
        

        Else
        
            OutputCellCount = OutputCellCount
            TotalShareVolume = ws.Cells((k + 1), 7).Value
            ws.Cells(OutputCellCount, 8) = TotalShareVolume

        End If

    Next k

'Finding the Total Share Volume
        
    OutputCellCount = 2
    TotalShareVolume = ws.Cells(2, 7).Value

    For q = 2 To x
        
        If ws.Cells(q + 1, 1).Value = ws.Cells(q, 1).Value Then
    
            TotalShareVolume = TotalShareVolume + ws.Cells(q + 1, 7).Value
     

        Else
        
            ws.Cells(OutputCellCount, 13) = TotalShareVolume
            OutputCellCount = OutputCellCount + 1
            TotalShareVolume = ws.Cells((q + 1), 7).Value

        End If

    Next q

    For r = 2 To y

        If ws.Cells(r, 11).Value > 0 Then
        
            ws.Cells(r, 10).Interior.ColorIndex = 4

        ElseIf ws.Cells(r, 11).Value < 0 Then
        
            ws.Cells(r, 10).Interior.ColorIndex = 3
            
        Else
        
            ws.Cells(r, 10).Interior.ColorIndex = 15
    
        End If

    Next r

'Bordering the data

    Set SummaryCells = ws.Range(ws.Cells(1, 10), ws.Cells(y, 14))

    SummaryCells.Borders.LineStyle = xlContinuous
    SummaryCells.Borders.Weight = 2

    Set AdditionalSummaryCells = ws.Range(ws.Cells(1, 16), ws.Cells(y, 17))

    AdditionalSummaryCells.Borders.LineStyle = xlContinuous
    AdditionalSummaryCells.Borders.Weight = 2

    Set AdditionalSummaryCells2 = ws.Range(ws.Cells(1, 19), ws.Cells(2, 20))

    AdditionalSummaryCells2.Borders.LineStyle = xlContinuous
    AdditionalSummaryCells2.Borders.Weight = 2

    Set SummaryCells = ws.Range(ws.Cells(7, 19), ws.Cells(10, 21))

    SummaryCells.Borders.LineStyle = xlContinuous
    SummaryCells.Borders.Weight = 2



'Changing Header Font Colour and Fill.

    ws.Range("A1:T1").Font.Bold = True
    ws.Range("J1:N1").Interior.ColorIndex = 16
    ws.Range("P1:Q1").Interior.ColorIndex = 16
    ws.Range("S1:T1").Interior.ColorIndex = 16
    ws.Range("T7:U7").Interior.ColorIndex = 16
    ws.Range("J1:N1").Font.ColorIndex = 2
    ws.Range("P1:Q1").Font.ColorIndex = 2
    ws.Range("S1:T1").Font.ColorIndex = 2
    ws.Range("T7:U7").Font.ColorIndex = 2
    

    ws.Range("S5").Font.Bold = True
    ws.Range("S5").Font.ColorIndex = 3
    ws.Range("T7:U7").Font.Bold = True
    ws.Range("T13:U13").Font.Bold = True
    
    'Bonus - Finding the Greatest Percentage Increase (For Year)

    LargerIncrease = ws.Cells(2, 12).Value
    ws.Range("T8").Value = ws.Cells(2, 12).Value
    ws.Range("U8").Value = ws.Cells(2, 10).Value

    For n = 2 To DifferentSharesCount

        If ws.Cells(n + 1, 12).Value > LargerIncrease Then

            LargerIncrease = ws.Cells(n + 1, 12).Value
            ws.Cells(8, 20).Value = LargerIncrease
            ws.Cells(8, 21) = ws.Cells(n + 1, 10)

        Else

            LargerIncrease = LargerIncrease

        End If
        
    Next n


'Bonus - Finding the Greatest Percentage Decrease (For Year)

    GreaterDecrease = ws.Cells(2, 12).Value
    ws.Range("T9").Value = ws.Cells(2, 12).Value
    ws.Range("U9").Value = ws.Cells(2, 10).Value

    For n = 2 To DifferentSharesCount

        If ws.Cells(n + 1, 12).Value <= GreaterDecrease Then

            GreaterDecrease = ws.Cells(n + 1, 12).Value
            ws.Cells(9, 20).Value = GreaterDecrease
            ws.Cells(9, 21).Value = ws.Cells(n + 1, 10).Value

        Else

            GreaterDecrease = GreaterDecrease

        End If
        
    Next n

'Bonus - Finding the Greatest Stock Volume (For Year)

    LargerVolume = ws.Cells(2, 13).Value
    ws.Range("T10").Value = ws.Cells(2, 12).Value
    ws.Range("U10").Value = ws.Cells(2, 10).Value

    For n = 2 To DifferentSharesCount

        If ws.Cells(n + 1, 13).Value > LargerVolume Then

            LargerVolume = ws.Cells(n + 1, 13).Value
            ws.Cells(10, 20).Value = LargerVolume
            ws.Cells(10, 21) = ws.Cells(n + 1, 10)

        Else

            LargerVolume = LargerVolume


        End If
        
    Next n
    
    Next ws
    
'Start processes that will only apply to the FIRST sheet.

'Variables that apply to FIRST sheet only (Not whole workbook).

    Dim WS_Count As Integer

    Dim f As Integer
    Dim m As Integer

    Dim BiggestOverallIncrease As Double
    Dim BiggestOverallDecrease As Double
    Dim OverallVolume As Double

    Dim a As Integer
    Dim b As Integer
    Dim c As Integer

 
'Bonus Table - All Years

    Range("T13").Value = "All Years:"
    Range("U13").Value = "Ticker:"

    Range("S14").Value = "Greatest % Increase:"
    Range("S15").Value = "Greatest % Decrease:"
    Range("S16").Value = "Greatest Stock Volume:"
    
    Range("T13:U13").Font.ColorIndex = 2
    Range("T13:U13").Interior.ColorIndex = 16
    
    
    Set SummaryCells = Range(Cells(13, 19), Cells(16, 21))
    
    SummaryCells.Borders.LineStyle = xlContinuous
    SummaryCells.Borders.Weight = 2
    
    'List of Leading Values from ALL sheets.
    
    Range("V1").Columns.ColumnWidth = 10
    
    Range("W1").Value = "Year"
    Range("W1").Columns.ColumnWidth = 6
    Range("X1").Value = "Greatest % Increase:"
    Range("X1").Columns.AutoFit
    Range("Y1").Value = "Ticker"
    Range("Y1").Columns.ColumnWidth = 6
    Range("W1:Y1").Font.ColorIndex = 2
    Range("W1:Y1").Interior.ColorIndex = 16
    
    Range("Z1").Columns.ColumnWidth = 10
    
    Range("AA1").Value = "Year"
    Range("AA1").Columns.ColumnWidth = 6
    Range("AB1").Value = "Greatest % Decrease:"
    Range("AB1").Columns.AutoFit
    Range("AC1").Value = "Ticker"
    Range("AC1").Columns.ColumnWidth = 6
    Range("AA1:AC1").Font.ColorIndex = 2
    Range("AA1:AC1").Interior.ColorIndex = 16
    
    Range("AD1").Columns.ColumnWidth = 10

    Range("AE1").Value = "Year"
    Range("AE1").Columns.ColumnWidth = 6
    Range("AF1").Value = "Greatest Stock Volume:"
    Range("AF1").Columns.AutoFit
    Range("AG1").Value = "Ticker"
    Range("AG1").Columns.ColumnWidth = 6
    Range("AE1:AG1").Font.ColorIndex = 2
    Range("AE1:AG1").Interior.ColorIndex = 16
    
    
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    f = 1
    
    For m = 1 To WS_Count
    
    ActiveWorkbook.Worksheets(1).Cells(f + 1, 24).Value = ActiveWorkbook.Worksheets(m).Range("T8").Value
    ActiveWorkbook.Worksheets(1).Cells(f + 1, 25).Value = ActiveWorkbook.Worksheets(m).Range("U8").Value
    ActiveWorkbook.Worksheets(1).Cells(f + 1, 28).Value = ActiveWorkbook.Worksheets(m).Range("T9").Value
    ActiveWorkbook.Worksheets(1).Cells(f + 1, 29).Value = ActiveWorkbook.Worksheets(m).Range("U9").Value
    ActiveWorkbook.Worksheets(1).Cells(f + 1, 32).Value = ActiveWorkbook.Worksheets(m).Range("T10").Value
    ActiveWorkbook.Worksheets(1).Cells(f + 1, 33).Value = ActiveWorkbook.Worksheets(m).Range("U10").Value
    
    Cells(f + 1, 23).Value = m + 2017
    Cells(f + 1, 27).Value = m + 2017
    Cells(f + 1, 31).Value = m + 2017
    
    f = f + 1
    
    Next m
    
    Set SummaryCells = Range(Cells(1, 23), Cells(WS_Count + 1, 25))
    
    SummaryCells.Borders.LineStyle = xlContinuous
    SummaryCells.Borders.Weight = 2
    
    Set SummaryCells = Range(Cells(1, 27), Cells(WS_Count + 1, 29))
    
    SummaryCells.Borders.LineStyle = xlContinuous
    SummaryCells.Borders.Weight = 2
    
    Set SummaryCells = Range(Cells(1, 31), Cells(WS_Count + 1, 33))
    
    SummaryCells.Borders.LineStyle = xlContinuous
    SummaryCells.Borders.Weight = 2
    
'Bonus - Finding the Greatest Percentage Increase (Overall)

    BiggestOverallIncrease = Range("X2").Value
    Range("T14").Value = Range("X2").Value
    Range("U14").Value = Range("Y2").Value
    
    For a = 2 To WS_Count

        If Cells(a + 1, 24).Value >= BiggestOverallIncrease Then

            BiggestOverallIncrease = Cells(a + 1, 24).Value
            Cells(14, 20).Value = BiggestOverallIncrease
            Cells(14, 21).Value = Cells(a + 1, 25).Value

        Else

            BiggestOverallIncrease = BiggestOverallIncrease

        End If
        
    Next a
    
'Bonus - Finding the Greatest Percentage Decrease (Overall)

    BiggestOverallDecrease = Range("AB2").Value
    Range("T15").Value = Range("AB2").Value
    Range("U15").Value = Range("AC2").Value

    For b = 2 To WS_Count

        If Cells(b + 1, 28).Value <= BiggestOverallDecrease Then

            BiggestOverallDecrease = Cells(b + 1, 28).Value
            Cells(15, 20).Value = BiggestOverallDecrease
            Cells(15, 21).Value = Cells(b + 1, 29).Value

        Else

            BiggestOverallDecrease = BiggestOverallDecrease
            Cells(15, 20).Value = BiggestOverallDecrease
            Cells(15, 21).Value = Range("AC2").Value


        End If
        
    Next b
    
    'Bonus - Finding the Greatest Stock Volume (Overall)

    OverallVolume = Range("AF2").Value
    Range("T16").Value = Range("AF2").Value
    Range("U16").Value = Range("AG2").Value
    
    For c = 2 To WS_Count

        If Cells(c + 1, 32).Value >= OverallVolume Then

            OverallVolume = Cells(c + 1, 32).Value
            Cells(14, 20).Value = OverallVolume
            Cells(14, 21).Value = Cells(c + 1, 33).Value

        Else

            OverallVolume = OverallVolume
            Cells(16, 20).Value = OverallVolume
            Cells(16, 21).Value = Range("AG2").Value

        End If
        
    Next c
        
End Sub



