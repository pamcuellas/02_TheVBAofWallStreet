    Option Explicit
    Sub mainCode()
    ' This is the main sub. Execute it to compute the numbers.'    
        Dim lastRowSheet As Long
        Dim colSource As String
        Dim colDate As String
        Dim colOpen As String
        Dim colClose As String
        Dim colVol As String
        Dim colNewTickers As String
        Dim colSTV As String
        Dim colYearly As String
        Dim colPChange As String
        Dim lastDate As String
        Dim strAux As String
        Dim lastValue As Double
        Dim firstValue As Double
        Dim lastRow As String
        Dim firstRow As String
        Dim totalSheets As Integer
        Dim red As Long
        Dim green As Long
        Dim currRow, currSheet As Integer
        Dim lastRowDistinctTickers As Integer
        
        red = 5263615
        green = 5287936
            
        colSource = "A"
        colDate = "B"
        colOpen = "C"
        colClose = "F"
        colVol = "G"
        colNewTickers = "I"
        colYearly = "J"
        colPChange = "K"
        colSTV = "L"
        lastRowSheet = Cells(Rows.Count, colSource).End(xlUp).Row
        totalSheets = Worksheets.Count
        

        'Loop for each Sheet.
        For currSheet = 1 To totalSheets Step 1
        
            Worksheets(currSheet).Select
            
            ' Create a distinct tickers list for the current Sheet.
            getDistinctTickers lastRowSheet, colSource, colNewTickers
            range(colNewTickers + "1").Value = "Ticker"
            
            'Get the last Date for the current Sheet. It is supposed to be the same for all Tickers.
            strAux = "=MAX(IF(" + colSource + "2:" + colSource & lastRowSheet & "=""" + _
                range(colNewTickers & 2).Value + """," + _
                colDate + "2:" + colDate & lastRowSheet & "), 2)"
            lastDate = Evaluate(strAux)
            
            ' Loop to compute numbers for each ticker.
            lastRowDistinctTickers = Cells(Rows.Count, colNewTickers).End(xlUp).Row
            For currRow = 2 To lastRowDistinctTickers
            
                'Get the first row for the current Ticker
                firstRow = "=MATCH(""" + range(colNewTickers & currRow).Value & """,A:A,0)"
                firstRow = Evaluate(firstRow)
            
                'Get the last row for the current Ticker
                lastRow = "=(MATCH(""" + range(colNewTickers & currRow).Value & """,A:A,0) + " + _
                "(COUNTIF(A:A,""" + range(colNewTickers & currRow).Value & """))-1)"
                lastRow = Evaluate(lastRow)
            
            
                '************** Yearly Change  **************
                ' Get the last close value for the current Ticker.
                lastValue = CDbl(range(colClose & lastRow).Value)
                
                ' Get the first open value for the current Ticker
                firstValue = CDbl(range(colOpen & firstRow).Value)
                If (firstValue = 0) Then
                    ' The opening price at the beginning of year is zero for the current Ticker.
                    ' So, we will try get the first price different from zero.
                    firstValue = getNextOpenValue(colOpen, firstRow, lastRow)
                    If firstValue = 0 Then
                        ' It was not possible. We will prevent division by zero.
                        If lastValue > 0 Then
                            firstValue = lastValue
                        Else
                            firstValue = 0.00000001
                        End If
                    End If
                End If
                
                ' Compute the Yearly Change for the current Ticker.
                range(colYearly & currRow).Value = lastValue - firstValue
                If range(colYearly & currRow).Value >= 0 Then
                    setColor (colYearly & currRow), green
                Else
                    setColor (colYearly & currRow), red
                End If
                
                '************** Total Stock Volume  **************
                ' Compute the Total Stock Volume for the current Ticker.
                range(colSTV & currRow).Value = Application.SumIf(range(colSource & firstRow & ":" + colSource & lastRow), _
                range(colNewTickers & currRow), range(colVol & firstRow & ":" + colVol & lastRow))
                
                
                '************** Percent Change  **************
                ' Compute the Percent Change for the current Ticker.
                range(colPChange & currRow).Value = range(colYearly & currRow).Value / firstValue
                
            Next
            ' Create a title for new columns.
            range(colYearly + "1").Value = "Early Change"
            range(colPChange + "1").Value = "Percent Change"
            range(colSTV + "1").Value = "Stock Total Volume"
            
            ' The Greatest results.
            range("N1").Value = "The Greatest results"
            range("N2").Value = "Greatest % Increase"
            range("N3").Value = "Greatest % Decrease"
            range("N4").Value = "Gretest Total Volume"
            range("P1").Value = "Value"
            range("P2").Value = Application.Max(range("K2:K" & lastRowDistinctTickers))
            range("P3").Value = Application.Min(range("K2:K" & lastRowDistinctTickers))
            range("P4").Value = Application.Max(range("L2:L" & lastRowDistinctTickers))
            range("O1").Value = "Ticher"
            strAux = Application.Match(range("P2"), range("K1:K" & lastRowDistinctTickers), 0)
            range("O2").Value = range(colNewTickers & strAux).Value
            strAux = Application.Match(range("P3"), range("K1:K" & lastRowDistinctTickers), 0)
            range("O3").Value = range(colNewTickers & strAux).Value
            strAux = Application.Match(range("P4"), range("L1:L" & lastRowDistinctTickers), 0)
            range("O4").Value = range(colNewTickers & strAux).Value
            
            ' Formatting numbers.
            formatRightZeros (colYearly & "2")
            setPercentFormat (colPChange & "2")
            range("P2:P3").Select
            selection.Style = "Percent"
            selection.NumberFormat = "0.00%"
            Columns("I:P").EntireColumn.AutoFit
        Next
        
        ' Back to the first sheet
        Worksheets(1).Select
        
    End Sub
    Sub getDistinctTickers(lastRow, colSource, colNewTickers)
        ActiveSheet.range(colSource + "1:" + colSource & lastRow).AdvancedFilter _
        Action:=xlFilterCopy, CopyToRange:=ActiveSheet.range(colNewTickers + "1"), Unique:=True
    End Sub
    Function getNextOpenValue(column, firstRow, lastRow) As Double
        Dim currCell As range
        getNextOpenValue = 0
        For Each currCell In range("$" + column + "$" & firstRow, "$" + column + "$" & lastRow)
          If CLng(currCell.Value) > 0 Then
             getNextOpenValue = currCell.Value
             Exit For
          End If
        Next currCell
    End Function
    Sub setColor(pRange, color)
        range(pRange).Select
        With selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .color = color
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End Sub
    Sub formatRightZeros(pRange)
        range(pRange).Select
        range(selection, selection.End(xlDown)).Select
        selection.NumberFormat = "0.000000000"
    End Sub
    Sub setPercentFormat(pRange)
        range(pRange).Select
        range(selection, selection.End(xlDown)).Select
        selection.Style = "Percent"
        selection.NumberFormat = "0.00%"
    End Sub
    Sub cleanSheets()
    ' Execute this Sub to clean the Sheets and then you can run Sub mainCode again.
        Dim currSheet As Integer

        For currSheet = 1 To Worksheets.Count Step 1
            Worksheets(currSheet).Select
            range("I1:P1").Select
            range(selection, selection.End(xlDown)).Select
            selection.ClearContents
        Next
        ' Back to the first sheet
        Worksheets(1).Select
    End Sub
