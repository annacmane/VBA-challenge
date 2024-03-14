Attribute VB_Name = "Module1"
Sub vbachallenge()

For Each ws In Worksheets
    Set ws = ThisWorkbook.Sheets("2018")
    
    Dim openprice As Double
    Dim closeprice As Double
    Dim stockname As String
    Dim stocktotal As Double
    Dim yearchange As Double
    Dim percentchange As Double
    Dim startofYear As Date
    Dim endofYear As Date
    Dim lastrow As Double
    Dim summarytable As Double
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    summarytablerow = 2
    stocktotal = 0
    
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            stockname = ws.Cells(i, 1).Value
            stocktotal = stocktotal + ws.Cells(i, 7).Value
            openprice = 0
            closeprice = 0
            
            For j = 2 To lastrow
                If ws.Cells(j, 1).Value = stockname Then
                    If ws.Cells(j, 2).Value = "20180102" Or "20190102" Or "20200102" Then
                        openprice = ws.Cells(j, 3).Value
                    ElseIf ws.Cells(j, 2).Value = "20181231" Or "20191231" Or "20201231" Then
                        closeprice = ws.Cells(j, 6).Value
                    End If
                End If
            Next j
        yearchange = closepirce - openprice
        If openprice <> 0 Then
            percentchange = (yearchange / openprice) * 100
        Else
            percentchange = 0
        End If
        
        ws.Range("I" & summarytable).Value = stockname
        ws.Range("J" & summarytable).Value = yearchange
        ws.Range("K" & summarytable).Value = percentchange
        ws.Range("L" & summarytable).Value = stocktotal
        
        summarytable = summarytable + 1
        stocktotal = 0
        End If
    Next i
    
    For k = 2 To lastrow
        If ws.Range("I" & k).Value > 0 Then
            ws.Range("I" & k).Interior.ColorIndex = 4
        If ws.Range("I" + k).Value = 0 Then
            ws.Range("I" & k).Interior.ColorIndex = 15
        If ws.Range("I" & k).Value < 0 Then
            ws.Range("I" & k).Interior.ColorIndex = 3
        End If
    Next k
    
    Dim minPercent As Double
    Dim maxPercent As Double
    Dim maxVolume As Double
    
    For l = 2 To lastrow
        If minPercent = ws.WorksheetFunction.Min(Range("I" & l)) Then
            stockname = ws.Range("O3").Value
            minPercent = ws.Range("P3").Value
        
        If maxPercent = ws.WorksheetFunction.Max(Range("I" & l)) Then
            stockname = ws.Range("O2").Value
            minPercent = ws.Range("P2").Value
        
        If maxVolume = ws.WorksheetFunction.Max(Range("L" & l)) Then
            stockname = ws.Range("O4").Value
            minPercent = ws.Range("P4").Value
        End If
    Next l
    ' name new cells
        ws.Range("I1,O1").Value = "TICKER"
        ws.Range("J1").Value = "YEAR CHANGE"
        ws.Range("K1").Value = "PERCENT CHANGE"
        ws.Range("L1").Value = "TOTAL STOCK VOLUME"
        ws.Range("N2").Value = "GREATAEST % INCREASE"
        ws.Range("N3").Value = "GREATEST% DECREASE"
        ws.Range("N4").Value = "GREATEST STOCK VOLUME"
        ws.Range("P1").Value = "VALUE"
        
        
        
Next ws
    
End Sub
