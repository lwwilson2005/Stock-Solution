Sub StockSolution()
'*****************************
'written by LWilson
'
'12/14/2017
'logic the old fashion way
'Assumptions
'1. I was told to use the .csv file in class
'       This file was a flat file with no multiple worksheets
'       If I had to process multiple worksheets I would have aggregated all of the worksheets
'           into one worksheet as we did in class. This would be my starting point for this code.
'*****************************
'brand = ticker
'tchg = total change
'sval = starting value for ticker
'eval = ending value for ticker
'bmax = ticker max closing value
'bmin = ticker minimum closing value
'bvol = ticker total volume
'pcchg = percent change for ticker
'avgdchg = average daily change
'periodcount = the number of periods counted for ticker
'dchg = daily change
'"l" variables are local variables only
    Dim i As Long, j As Long
    Dim brand(20000) As String, tChg(20000) As Double, sval(20000) As Double, eval(20000) As Double
    Dim bMax(20000) As Double, bMin(20000) As Double, bVol(20000) As Double
    Dim pcChg(20000) As Double, avgDChg(20000) As Double, periodCount(20000) As Long, dChg(20000) As Double
    Dim lMin As Double, lMax As Double, lVol As Double, lClose As Double
    Dim lrowcount As Long
    Dim createNewSheet As Boolean  'this will determine if we create the sheet or continue to test
    Dim multipleDataWorksheets As Boolean   'this is used if the data comes from multiple worksheets
    createNewSheet = True
    multipleDataWorksheets = False
    j = 1       'ticker index
    lrowcount = ActiveSheet.UsedRange.Rows.Count    'last row count
    'next seed the starting values
    lMax = Round(Cells(2, 6).Value, 2)
    lMin = Round(Cells(2, 6).Value, 2)
    lVol = Cells(2, 7)
    sval(j) = Round(Cells(2, 6).Value, 2)  'starting value of first ticker
    k = 1       'count number of periods
    '
    'main loop to gather data
    For i = 2 To lrowcount
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then  'last ticker
            'Debug.Print i, Cells(i, 1).Value
            lVol = Cells(i, 7).Value
            lClose = Round(Cells(i, 6).Value, 2)
            brand(j) = Cells(i, 1).Value
            'tChg(j) = tChg(j) + lClose
            bVol(j) = bVol(j) + lVol
            eval(j) = lClose       'ending value of ticker
            bMin(j) = lMin
            bMax(j) = lMax
            periodCount(j) = k
            If i = 2 Then
                dChg(j) = 0
            Else
                dChg(j) = dChg(j) + Round(Abs(Cells(i, 4).Value - Cells(i, 5).Value), 2) ' high - low
            End If
            'cleanup and start next ticker
            j = j + 1
            k = 1
            sval(j) = Cells(i + 1, 6)   'starting value of next ticker
            lMax = Cells(i + 1, 6)
            lMin = Cells(i + 1, 6)
            
            
        Else
            lVol = Cells(i, 7).Value
            lClose = Round(Cells(i, 6).Value, 2)
            If lClose < lMin Then       'store new lmin
                lMin = lClose
            End If
            If lClose > lMax Then       'store new lmax
                lMax = lClose
            End If
            If i = 2 Then
                dChg(j) = 0
            Else
                dChg(j) = dChg(j) + Round(Abs(Cells(i, 4).Value - Cells(i, 5).Value), 2)  'high - low
            End If
            brand(j) = Cells(i, 1).Value
            'tChg(j) = tChg(j) + Cells(i, 3)
            bVol(j) = bVol(j) + lVol
            k = k + 1
        End If
        
    Next i
    'end of main loop
    
    'summarize results
    'Debug.Print j
    Dim GrVol As Double, GrIncr As Double, GrDecr As Double, GrAvgChg As Double
    'GrVol - greatest volume, index is grvolndx
    'GrIncr - greatest increase, index is GrIncrndx
    'GrDecr - greatest decrease, index is GrDecrndx
    'GrAvgChg - greatest average change, index is GrAvgChgndx
    GrVol = 0
    GrIncr = 0
    GrDecr = 0
    gravgdchg = 0
    For i = 1 To j - 1
        'loop through indexes to get greatest values
        'round for user readability
        
        tChg(i) = Round(eval(i) - sval(i), 2)
        pcChg(i) = Round((eval(i) - sval(i)) / sval(i), 4)
        avgDChg(i) = Round(dChg(i) / periodCount(i), 2)
        'Debug.Print brand(i), tChg(i), pcChg(i) * 100, avgDChg(i), bVol(i)
        'bMin(i), bMax(i), sval(i), eval(i), periodCount(i), dChg(i)
        If bVol(i) > GrVol Then
            GrVol = bVol(i)
            grvolndx = i
        End If
        If pcChg(i) > GrIncr Then
            GrIncr = pcChg(i)
            Grincrndx = i
        End If
        If pcChg(i) < GrDecr Then
            GrDecr = pcChg(i)
            Grdecrndx = i
        End If
        If avgDChg(i) > GrAvgChg Then
            GrAvgChg = avgDChg(i)
            GravgDChgndx = i
            Debug.Print gravgdchg, avgDChg(i), GrAvgChg
        End If
        
    Next i
    Debug.Print "greatest v = ", GrVol, grvolndx, brand(grvolndx)
    Debug.Print "greatest % incr = ", GrIncr, Grincrndx, brand(Grincrndx), sval(Grincrndx), eval(Grincrndx)
    Debug.Print "greatest % decr = ", GrDecr, Grdecrndx, brand(Grdecrndx), sval(Grdecrndx), eval(Grdecrndx)
    Debug.Print "greatest avgdchg = ", GrAvgChg, GravgDChgndx, brand(GravgDChgndx)
    
    'write new sheet and finish
    If createNewSheet Then
        Worksheets.Add      'adds worksheet to front and makes activesheet
        ActiveSheet.Name = "Summary"    'name the tab summary
        'add headers
        Range("a1").Value = "Ticker"
        Range("b1").Value = "Total Change"
        Range("c1").Value = "% Change"
        Range("d1").Value = "Avg. Daily Change"
        Range("e1").Value = "Volume"
        Range("g2").Value = "Greatest Volume"
        Range("g5").Value = "Greatest % Increase"
        Range("g8").Value = "Greatest % Decrease"
        Range("g11").Value = "Greatest Avg Change"
        Range("h5").NumberFormat = "0.00%"
        Range("h8").NumberFormat = "0.00%"
        Range("h11").NumberFormat = "0.00"
        'populate the sheet with data
        For i = 1 To j - 1
            Cells(i + 1, 1).Value = brand(i)
            Cells(i + 1, 2).Value = tChg(i)
            Cells(i + 1, 3).Value = pcChg(i)
            Cells(i + 1, 4).Value = avgDChg(i)
            Cells(i + 1, 5).Value = bVol(i)
            Cells(i + 1, 3).NumberFormat = "0.00%"
            If Cells(i + 1, 2) > 0 Then
                Cells(i + 1, 2).Interior.Color = vbGreen
            End If
            If Cells(i + 1, 2) < 0 Then
                Cells(i + 1, 2).Interior.Color = vbRed
            End If
        Next i
        Range("h2").Value = GrVol
        Range("I2").Value = brand(grvolndx)
        Range("h5").Value = GrIncr
        Range("I5").Value = brand(Grincrndx)
        Range("h8").Value = GrDecr
        Range("I8").Value = brand(Grdecrndx)
        Range("h11").Value = GrAvgChg
        Range("I11").Value = brand(GravgDChgndx)
        
    End If

End Sub

