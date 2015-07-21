Sub USA_Full_Data()
'
' USA Full Data Macro (Yearly)
'

Application.DisplayAlerts = False

Dim goCheck As Integer
Dim mySeries As Series
Dim seriesCol As SeriesCollection
Dim impexp As String
Dim inout As String

goCheck = MsgBox("This macro is designed for use with the raw data output (.csv) file from USA Trade Online (usatrade.census.gov)." & vbNewLine & vbNewLine & "For best results, use the " & Chr(34) & "USA Full Data (Macro-Ready)" & Chr(34) & " report template on USA Trade website, and select at least one full year and all sub-month data available." & vbNewLine & vbNewLine & "Do you want to proceed?", vbYesNoCancel)

If goCheck = vbYes Then

        If InStr(1, Range("A1").Text, "Import") Then
            impexp = "Import"
            inout = "to"
        ElseIf InStr(1, Range("A1").Text, "Export") Then
            impexp = "Export"
            inout = "from"
        End If
    
        Rows("1:1").Delete Shift:=xlUp
        Range("B2").Cut Destination:=Range("A2")
        Range("A2").Font.Bold = True
        Range("A5").ClearContents
        
        ActiveSheet.Name = "Main"
        
    LastHorzCol = Cells(5, Columns.Count).End(xlToLeft).Column
    
    'Merge and Create borders for each month/year
    For i = 0 To ((LastHorzCol - 1) / 3)
        j = ((3 * i) + 2)
        Columns(j).Select
        
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        Selection.Borders(xlEdgeTop).LineStyle = xlNone
        Selection.Borders(xlEdgeBottom).LineStyle = xlNone
        Selection.Borders(xlEdgeRight).LineStyle = xlNone
        Selection.Borders(xlInsideVertical).LineStyle = xlNone
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
        If j < LastHorzCol Then
           'Merge Countries
            Range(Cells(4, j), Cells(4, (j + 2))).Select
            With Selection
                .HorizontalAlignment = xlCenter
                .Merge
            End With
        End If
      
    Next i
    
        'Make horizontal line
        Rows("3:3").Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        Selection.Borders(xlEdgeLeft).LineStyle = xlNone
        Selection.Borders(xlEdgeTop).LineStyle = xlNone
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        Selection.Borders(xlEdgeRight).LineStyle = xlNone
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        
        'Delete extra useless line
        Range("A6").Select
        Selection.Cut Destination:=Range("A5")
        Rows("6:6").Select
        Selection.Delete Shift:=xlUp
        
        'Color District line grey.
        Range(Cells(5, 1), Cells(5, LastHorzCol)).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.249977111117893
            .PatternTintAndShade = 0
        End With
        
        'Color "All Districts / Totals" line darker grey.
        Range(Cells(6, 1), Cells(6, LastHorzCol)).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.349986266670736
            .PatternTintAndShade = 0
        End With
        
        
        'Rows("5:5").Select
        'Selection.Font.Bold = True
        Range(Columns(2), Columns(LastHorzCol)).EntireColumn.AutoFit
        Columns("A:A").ColumnWidth = 21
        
        'Find last row
        LastVertRow = Cells(Rows.Count, "A").End(xlUp).Row
        
        'Delete "Row Suppresion Applied."
        If InStr(1, Cells(LastVertRow, 1).Value, "Row suppression applied.") Then
            Range(Rows(LastVertRow), Rows((LastVertRow - 1))).Select
            Selection.Delete Shift:=xlUp
        End If
        If InStr(1, Cells(LastVertRow, 1).Value, "Row and column suppression applied.") Then
            Range(Rows(LastVertRow), Rows((LastVertRow - 1))).Select
            Selection.Delete Shift:=xlUp
        End If
        
        'Add another horizontal line
        Rows("3:3").Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        Selection.Borders(xlEdgeLeft).LineStyle = xlNone
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        Selection.Borders(xlEdgeRight).LineStyle = xlNone
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        
    For UnitPrice = 1 To ((LastHorzCol - 1) / 3)
        upr = ((3 * UnitPrice) + 1)
        Cells(6, upr).Select
        ActiveCell.FormulaR1C1 = "=IFERROR(RC[-2]/RC[-1], """")"
        Selection.AutoFill Destination:=Range(Cells(6, upr), Cells(LastVertRow, upr)), Type:=xlFillValues
        Range(Cells(6, upr), Cells(LastVertRow, upr)).Select
        Selection.NumberFormat = "0.000"
    Next UnitPrice
        
        Range("B6").Select
        ActiveWindow.FreezePanes = True
        
    'Sort country columns by quantity (highest first)
        Range("B3").Select
        
        Dim dateflag As Integer
        dateflag = 0
        
    For i = 2 To (LastHorzCol - 1)
        num_country = 0
        minuscountry = 0
        If (IsDate(Cells(3, i).Value) = False) And (Cells(4, i).Value = "World Total") Then
        ' Command goes here if it has found a year group
            ' Count number of countries in current year
            A = i + 3
            Do While (Cells(4, A).Value <> "World Total" And Cells(4, A).Value <> "")
                num_country = num_country + 1
                A = A + 3
            Loop
            
            'Now we have # countries, sort them by quantity:
            placeholder = i + 4
            For XX = 1 To num_country
                
                    For YY = 1 To (num_country - 1)
                        If Cells(6, placeholder).Value < Cells(6, (placeholder + 3)).Value Then
                            Range(Columns(placeholder - 1), Columns(placeholder + 1)).Select
                            Selection.Cut
                            Columns(placeholder + 5).Select
                            Selection.Insert Shift:=xlToRight
                        End If
                        placeholder = placeholder + 3
                    Next YY
                placeholder = i + 4
            Next XX
            
            'Remove columns with tiny quantities (under 18,000)
            
            For ZZ = 1 To num_country
                If Cells(5, placeholder).Value = "" Then
                
                ElseIf Cells(6, placeholder).Value < 18000 Then
                    Range(Columns(placeholder - 1), Columns(placeholder + 1)).Select
                    Selection.Delete
                    placeholder = placeholder - 3
                    minuscountry = minuscountry + 1
                End If
                placeholder = placeholder + 3
            Next ZZ
    
            If dateflag = 0 Then
                printrange = (num_country - minuscountry + 1) * 3 + 1
                Cells(1, printrange).Value = Cells(3, i).Value
                dateflag = 1
            End If
            
            i = i + 2
        ElseIf (IsDate(Cells(3, i).Value = True) And (Cells(4, i).Value = "World Total")) Then
            ' Command goes here to process months
            
            i = i + 2
        Else
            ' Do nothing if not a month or year
            i = i + 2
        End If
        
    Next i
    
    'Refresh column count
    LastHorzCol = Cells(5, Columns.Count).End(xlToLeft).Column
    
    'Merge months or years
    Range("H1").Select
    
    county = 0
    
    For ymm = 2 To LastHorzCol
        Cells(3, ymm).Select
        If ActiveCell.Value = "" Then
            Range(Cells(3, ymm), Cells(3, ymm - 1)).Select
            With Selection
                .HorizontalAlignment = xlCenter
                .Merge
            End With
        Else
            county = county + 1
            If county Mod 2 = 1 Then
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = -0.149998474074526
                    .PatternTintAndShade = 0
                End With
                Selection.Font.Bold = True
            Else
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = -0.499984740745262
                    .PatternTintAndShade = 0
                End With
                Selection.Font.Bold = True
            End If
        End If
    Next ymm
    
    For zmm = 7 To (LastVertRow - 2)
        
        Rows(zmm & ":" & zmm).Select
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.249946592608417
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.249946592608417
            .Weight = xlThin
        End With
        
    Next zmm
    
    Range("H1").Select
    
        'Create Chart Data page
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Graph Data"
        Sheets("Main").Activate
        Sheets("Graph Data").Range("AO30").Value = " "
        
    cyear = 0
    cmonth = 0
    ccountry = 2
    ci = 2
    cnummonths = 0
    flag = 0
    
    Do While ci <= LastHorzCol
        If IsDate(Sheets("Main").Cells(3, ci).MergeArea.Cells(1, 1).Value) = False Then
            'If true, we are in a year heading
         
            If flag = 0 Then
                cc = ci + 3
                dmonth = cmonth
                Do Until (Sheets("Main").Cells(3, cc).MergeArea.Cells(1, 1).Value = "" Or IsDate(Sheets("Main").Cells(3, cc).MergeArea.Cells(1, 1).Value) = False) And (Sheets("Main").Cells(3, cc - 3).MergeArea.Cells(1, 1).Value <> Sheets("Main").Cells(3, cc).MergeArea.Cells(1, 1).Value)
                    'Do until reach a new year
                    'Possible: Do until reach new month?
                    If (IsDate(Sheets("Main").Cells(3, cc).MergeArea.Cells(1, 1).Value) = True) And (Sheets("Main").Cells(3, cc - 1).MergeArea.Cells(1, 1).Value <> Sheets("Main").Cells(3, cc).MergeArea.Cells(1, 1).Value) Then
                        cnummonths = cnummonths + 1
                        'Add Month Headings
                        Sheets("Graph Data").Cells(((3 + cmonth) + (cyear * 30)), 1).Value = Sheets("Main").Cells(3, cc).MergeArea.Cells(1, 1).Value
                        Sheets("Graph Data").Cells(((18 + cmonth) + (cyear * 30)), 1).Value = Sheets("Main").Cells(3, cc).MergeArea.Cells(1, 1).Value
                        cmonth = cmonth + 1
                    End If
                    cc = cc + 3
                Loop
                flag = 1
                cmonth = dmonth
                'Restore month counter to previous
            End If
            
            ca = ci + 3
            cb = ca + 6
            
            If Sheets("Main").Cells(4, ca).Value <> "World Total" Then
                For c_x = 1 To cnummonths
                    Do Until (Sheets("Main").Cells(3, cb).MergeArea.Cells(1, 1).Value <> Sheets("Main").Cells(3, cb - 3).MergeArea.Cells(1, 1).Value) And (IsDate(Sheets("Main").Cells(3, cb - 3).MergeArea.Cells(1, 1).Value) = True Or IsDate(Sheets("Main").Cells(3, cb).MergeArea.Cells(1, 1).Value) = False)
                    'Do for one month
                        If Sheets("Main").Cells(4, ca).Value = Sheets("Main").Cells(4, cb).Value Then
                            
                            'Add Unit Price and Quantity Values
                            Sheets("Graph Data").Cells(((3 + cmonth) + (cyear * 30)), ccountry).Value = Sheets("Main").Cells(6, cb + 2).Value
                            Sheets("Graph Data").Cells(((18 + cmonth) + (cyear * 30)), ccountry).Value = Sheets("Main").Cells(6, cb + 1).Value
                                     
                            cmonth = cmonth + 1
                        End If
                        cb = cb + 3
                    Loop
                    If Sheets("Graph Data").Cells(((18 + cmonth - 1) + (cyear * 30)), ccountry).Value = "" Then
                        Sheets("Graph Data").Cells(((18 + cmonth - 1) + (cyear * 30)), ccountry).Value = ""
                        cmonth = cmonth + 1
                    End If
                    cb = cb + 3
                Next c_x
           
                'Add Country names
                Sheets("Graph Data").Cells(2 + (cyear * 30), ccountry).Value = Sheets("Main").Cells(4, ca).Value
                Sheets("Graph Data").Cells(17 + (cyear * 30), ccountry).Value = Sheets("Main").Cells(4, ca).Value
                'Add Header
                Sheets("Graph Data").Cells((1 + (cyear * 30)), 1).Value = "Unit Price"
                Sheets("Graph Data").Cells((16 + (cyear * 30)), 1).Value = "Quantity"
                'Add Year
                Sheets("Graph Data").Cells((2 + (cyear * 30)), 1).Value = Sheets("Main").Cells(3, ci).MergeArea.Cells(1, 1).Value
                Sheets("Graph Data").Cells((17 + (cyear * 30)), 1).Value = Sheets("Main").Cells(3, ci).MergeArea.Cells(1, 1).Value
                
                ccountry = ccountry + 1
            End If
            
            'Check if it's a new year or not, add to year counter and set up 2nd page if it is:
            If IsDate(Sheets("Main").Cells(3, (ci + 3)).MergeArea.Cells(1, 1).Value) = False And ((IsDate(Sheets("Main").Cells(3, (ci + 6)).MergeArea.Cells(1, 1).Value) = True) Or (Sheets("Main").Cells(3, (ci + 6)).MergeArea.Cells(1, 1).Value = "")) Then
                
                'This is where to tell it to make the charts
                ' ==========================================
                
                'Formatting:
                Sheets("Graph Data").Activate
                'Zoom Out
                ActiveWindow.Zoom = 80
                
                'Change Quantity Format [add comma, no decimal]
                With Sheets("Graph Data").Range(Cells(18 + (cyear * 30), 2), Cells(17 + cnummonths + (cyear * 30), ccountry - 1))
                    .Style = "Comma"
                    .NumberFormat = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-"
                End With
                
                'Add table borders
                
                Sheets("Graph Data").Range(Cells(2 + (cyear * 30), 1), Cells(2 + cnummonths + (cyear * 30), ccountry - 1)).Select
                For bords = 1 To 2
                    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                    With Selection.Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Selection.Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Selection.Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Selection.Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Selection.Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Selection.Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Range(Cells(17 + (cyear * 30), 1), Cells(18 + cnummonths + (cyear * 30), ccountry - 1)).Select
                Next bords
                
                'Add vertical light grey background
                    With Range(Cells(3 + (cyear * 30), 1), Cells(2 + cnummonths + (cyear * 30), 1)).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorDark1
                        .TintAndShade = -0.149998474074526
                        .PatternTintAndShade = 0
                    End With
                    With Range(Cells(18 + (cyear * 30), 1), Cells(18 + cnummonths + (cyear * 30), 1)).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorDark1
                        .TintAndShade = -0.149998474074526
                        .PatternTintAndShade = 0
                    End With
                'Add horizontal light grey background
                    With Range(Cells(2 + (cyear * 30), 1), Cells(2 + (cyear * 30), ccountry - 1)).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorDark1
                        .TintAndShade = -0.149998474074526
                        .PatternTintAndShade = 0
                    End With
                    With Range(Cells(17 + (cyear * 30), 1), Cells(17 + (cyear * 30), ccountry - 1)).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorDark1
                        .TintAndShade = -0.149998474074526
                        .PatternTintAndShade = 0
                    End With
                Range(Cells(2 + (cyear * 30), 1), Cells(2 + (cyear * 30), 1)).Select
                For bordz = 1 To 2
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorDark1
                        .TintAndShade = -0.349986266670736
                        .PatternTintAndShade = 0
                    End With
                    Selection.Font.Bold = True
                    With Selection
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlBottom
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    Range(Cells(17 + (cyear * 30), 1), Cells(17 + (cyear * 30), 1)).Select
                Next bordz
                              
                
                '======== Add code to change Quantity from kg to MT if appropriate
                If Right(Sheets("Main").Range("A2").Value, 4) = "(kg)" Then
                    For h = (18 + (cyear * 30)) To (17 + cnummonths + (cyear * 30))
                        
                        For v = 2 To (ccountry - 1)
                            Cells(h, v).Select
                            If IsError(Cells(h, v)) Then
                                               
                            ElseIf Cells(h, v).Value = 0 Then
                                            
                            Else
                                Cells(h, v).Value = Cells(h, v).Value / 1000
                            End If
                    
                        Next
                    
                    Next
                End If
                '==================================
                
                'Change Unit Price format [add 0.000 decimal]
                With Sheets("Graph Data").Range(Cells(3 + (cyear * 30), 2), Cells(2 + cnummonths + (cyear * 30), ccountry - 1))
                    .NumberFormat = "0.000"
                End With
                'Change Month Listing Format
                'for Quantity
                With Sheets("Graph Data").Range(Cells(3 + (cyear * 30), 1), Cells(2 + cnummonths + (cyear * 30), 1))
                    .NumberFormat = "mmm"
                End With
                'for Unit Price
                With Sheets("Graph Data").Range(Cells(18 + (cyear * 30), 1), Cells(17 + cnummonths + (cyear * 30), 1))
                    .NumberFormat = "mmm"
                End With
                Columns("A:Z").EntireColumn.AutoFit
                           
                'Unit Price Chart
                ActiveSheet.Shapes.AddChart.Select
                With ActiveChart
                    .ChartType = xlLineMarkers
                    .SetSourceData Source:=Range(Cells(2 + (cyear * 30), 1), Cells(2 + cnummonths + (cyear * 30), ccountry - 1))
                    .Parent.Top = 0 + (cyear * 450)
                    .Parent.Left = Range("N1").Left
                    .Parent.Width = 485
                    .Parent.Height = 400
                    .SetElement (msoElementPrimaryValueAxisTitleRotated)
                    .Axes(xlValue, xlPrimary).AxisTitle.Text = "Unit Price ($/" & Left(Right(Sheets("Main").Range("A2").Value, 3), 2) & ")"
                    
                    .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
                    .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Month"
                    
                    .SetElement (msoElementChartTitleAboveChart)
                    .ChartTitle.Text = Cells(17 + (cyear * 30), 1) & " - " & Right(Sheets("Main").Range("A2").Value, (Len(Sheets("Main").Range("A2").Value) - 11)) & " Unit Price " & inout & " USA, Monthly"
                    
                    
                    'Change markers to little circles
                    seriescount = 1
                    Set seriesCol = ActiveChart.SeriesCollection
                    For Each mySeries In seriesCol
                        Set mySeries = ActiveChart.SeriesCollection(seriescount)
                        With mySeries
                            .MarkerStyle = 8
                            .MarkerSize = 5
                        End With
                        seriescount = seriescount + 1
                    Next
                    
                    .PlotArea.Width = 400
                    .PlotArea.Left = 70
                    'Add data table, move headings around
                    .SetElement (msoElementDataTableWithLegendKeys)
                    .Axes(xlCategory).Delete
                    .Legend.Delete
                    .DataTable.Font.Size = 9
                    .Axes(xlCategory).AxisTitle.Delete
                    .Axes(xlValue).AxisTitle.Left = 6
                    .Axes(xlValue).AxisTitle.Top = ActiveChart.PlotArea.Height / 3
                    
                    
                End With
                
                'Quantity Chart
                ActiveSheet.Shapes.AddChart.Select
                With ActiveChart
                    .ChartType = xlLineMarkers
                    .SetSourceData Source:=Range(Cells(17 + (cyear * 30), 1), Cells(17 + cnummonths + (cyear * 30), ccountry - 1))
                    .Parent.Top = 0 + (cyear * 450)
                    .Parent.Left = Range("Y1").Left
                    .Parent.Width = 485
                    .Parent.Height = 400
                                   
                    .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
                    .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Month"
                    
                    .SetElement (msoElementChartTitleAboveChart)
                    .ChartTitle.Text = Cells(17 + (cyear * 30), 1) & " - " & Left(Right(Sheets("Main").Range("A2").Value, (Len(Sheets("Main").Range("A2").Value) - 11)), (Len(Right(Sheets("Main").Range("A2").Value, (Len(Sheets("Main").Range("A2").Value) - 11))) - 5)) & " Quantity " & impexp & "ed (MT) " & inout & " USA, Monthly"
                    
                    .SetElement (msoElementPrimaryValueAxisTitleRotated)
                    
                    .Axes(xlValue, xlPrimary).AxisTitle.Text = "Quantity " & impexp & "ed (MT)"

                    'Change markers to little circles
                    seriescount = 1
                    Set seriesCol = ActiveChart.SeriesCollection
                    For Each mySeries In seriesCol
                        Set mySeries = ActiveChart.SeriesCollection(seriescount)
                        With mySeries
                            .MarkerStyle = 8
                            .MarkerSize = 5
                        End With
                        seriescount = seriescount + 1
                    Next
                    
                    .PlotArea.Width = 400
                    .PlotArea.Left = 70
                    'Add data table, move headings around
                    .SetElement (msoElementDataTableWithLegendKeys)
                    .Axes(xlCategory).Delete
                    .Legend.Delete
                    .DataTable.Font.Size = 9
                    .Axes(xlCategory).AxisTitle.Delete
                    .Axes(xlValue).AxisTitle.Left = 6
                    .Axes(xlValue).AxisTitle.Top = ActiveChart.PlotArea.Height / 3
                End With

                ' ==========================================
                Sheets("Main").Activate
                cyear = cyear + 1
                flag = 0
                cnummonths = 0
            End If
    
        Else
            ccountry = 2
        End If
        cmonth = 0
        If (IsDate(Sheets("Main").Cells(3, (ci + 6)).Value) = True) And (Sheets("Main").Cells(4, (ci + 6)).Value = "World Total") Then
            ci = ci + 3
        End If
        ci = ci + 3
    Loop
    
    '======== This is where year-total chart code goes ========='
    
    LastRow = Sheets("Graph Data").Range("A" & Rows.Count).End(xlUp).Row
 
    Sheets("Graph Data").Activate
    
    Tots = 1
    If cyear > 1 Then
        For Tots = 1 To (cyear - 1)
            moc = 2
            Cells(30 * Tots, 1).Value = "TOTAL"
            Do While Cells(17 + (30 * (Tots - 1)), moc).Value <> ""
                Cells(30 * Tots, moc).FormulaR1C1 = "=SUM(R[-12]C:R[-1]C)"
                moc = moc + 1
            Loop
        Next Tots
    End If
    'tots = tots + 1 is this neccesary?
    
    Cells(LastRow + cyear - (cyear - 1), 1).Value = "TOTAL"
    cmo = 0
    Do While IsDate(Cells(LastRow + cyear - (cyear - 1) - 1 - cmo, 1).Value) = True
        cmo = cmo + 1 'Count number of months in last year listed
    Loop
    moc = 2
    Do While Cells(17 + (30 * (Tots - 1)), moc).Value <> ""
        Cells(LastRow + cyear - (cyear - 1), moc).FormulaR1C1 = "=SUM(R[-" & cmo & "]C:R[-1]C)"
        moc = moc + 1
    Loop
    
    
    'Add spacer
    For Tots = 1 To cyear
        Rows((31 * Tots) & ":" & (31 * Tots)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        With Cells(31 * Tots, 1).Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    Next Tots
    
    ' ============= YEARLY TOTALS TABLE ==========
      
    LastRow = Sheets("Graph Data").Range("A" & Rows.Count).End(xlUp).Row
    
    
    If cyear > 1 Then
        'Create headers
            Cells((LastRow + 2), 1).Value = "Yearly Totals"
            
            colz = 2
            flaggyb = 0
            Do While Cells(17, colz).Value <> ""
                Cells((LastRow + 2), colz).Value = Cells(17, colz).Value
                colz = colz + 1
                flaggyb = flaggyb + 1
            Loop

        'Add data
        For Tots = 1 To (cyear - 1)
            Cells((LastRow + 2 + Tots), 1).Value = Cells((17 + 31 * (Tots - 1)), 1).Value
            colz = 2 'reference
            Do While Cells((17 + 31 * (Tots - 1)), colz).Value <> ""
                colb = 2 'moving in result header
                flaggy = 0
                Do While Cells((LastRow + 2), colb).Value <> ""
                    If Cells((LastRow + 2), colb).Value = Cells((17 + 31 * (Tots - 1)), colz).Value Then
                        Cells((LastRow + 2 + Tots), colb).Value = Cells((30 + 31 * (Tots - 1)), colz).Value 'Problem here?
                        flaggy = 1
                    End If
                    colb = colb + 1
                Loop
                               
                If flaggy = 0 Then
                    flaggyb = 1
                    Do While Cells((LastRow + 2), flaggyb).Value <> ""
                        flaggyb = flaggyb + 1
                    Loop
                    Cells((LastRow + 2), flaggyb).Value = Cells((17 + 31 * (Tots - 1)), colz).Value
                    Cells((LastRow + 2 + Tots), flaggyb).Value = Cells((30 + 31 * (Tots - 1)), colz).Value 'Problem here?
                End If
                colz = colz + 1
            Loop
        Next Tots
        
        Cells((LastRow + 2 + Tots), 1).Value = Cells((17 + 31 * (Tots - 1)), 1).Value
        
        colz = 2 'reference
        Do While Cells((17 + 31 * (Tots - 1)), colz).Value <> ""
            colb = 2 'moving in result header
            flaggy = 0
            Do While Cells((LastRow + 2), colb).Value <> ""
                If Cells((LastRow + 2), colb).Value = Cells((17 + 31 * (Tots - 1)), colz).Value Then
                    Cells((LastRow + 2 + Tots), colb).Value = Cells(LastRow, colz).Value
                    flaggy = 1
                End If
                colb = colb + 1
            Loop
                           
            If flaggy = 0 Then
                flaggyb = 1
                Do While Cells((LastRow + 2), flaggyb).Value <> ""
                    flaggyb = flaggyb + 1
                Loop
                Cells((LastRow + 2), flaggyb).Value = Cells((17 + 31 * (Tots - 1)), colz).Value
                Cells((LastRow + 2 + Tots), flaggyb).Value = Cells(LastRow, colz).Value
            End If
            colz = colz + 1
        Loop

        ' ======= END DATA, BEGIN FORMATTING ========
        
        With Range(Cells(LastRow + 3, 2), Cells(LastRow + 2 + cyear, flaggyb))
            .Style = "Comma"
            .NumberFormat = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-"
        End With
        With Range(Cells(LastRow + 2, 1), Cells(LastRow + 2 + cyear, flaggyb))
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
        'Add grey background vert
        With Range(Cells(LastRow + 3, 1), Cells(LastRow + 2 + cyear, 1)).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.149998474074526
            .PatternTintAndShade = 0
        End With
        'Add grey background horz
        With Range(Cells(LastRow + 2, 2), Cells(LastRow + 2, flaggyb)).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.149998474074526
            .PatternTintAndShade = 0
        End With
        
        'Add dark grey background single cell
        With Cells(LastRow + 2, 1)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.349986266670736
                .PatternTintAndShade = 0
            End With
        End With
        
        ' === Create chart for yearly totals ======
        ActiveSheet.Shapes.AddChart.Select
        With ActiveChart
            .ChartType = xlLineMarkers
            .SetSourceData Source:=Range(Cells(LastRow + 2, 1), Cells(LastRow + 2 + cyear, flaggyb))
            .Parent.Top = ((LastRow + cyear + 3) * 15)
            .Parent.Left = 10
            .Parent.Width = 450
            .Parent.Height = 400
                           
            .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Year"
            
            .SetElement (msoElementChartTitleAboveChart)
            .ChartTitle.Text = Cells(LastRow + 2, 1).Value & " - " & Left(Right(Sheets("Main").Range("A2").Value, (Len(Sheets("Main").Range("A2").Value) - 11)), (Len(Right(Sheets("Main").Range("A2").Value, (Len(Sheets("Main").Range("A2").Value) - 11))) - 5)) & " Quantity " & impexp & "ed (MT) " & inout & " USA, Yearly"
            
            .SetElement (msoElementPrimaryValueAxisTitleRotated)
            
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "Quantity " & impexp & "ed (MT)"

            'Change markers to little circles
            seriescount = 1
            Set seriesCol = ActiveChart.SeriesCollection
            For Each mySeries In seriesCol
                Set mySeries = ActiveChart.SeriesCollection(seriescount)
                With mySeries
                    .MarkerStyle = 8
                    .MarkerSize = 5
                End With
                seriescount = seriescount + 1
            Next
            
            .PlotArea.Width = 400
            .PlotArea.Left = 70
            'Add data table, move headings around
            .SetElement (msoElementDataTableWithLegendKeys)
            .Axes(xlCategory).Delete
            .Legend.Delete
            .DataTable.Font.Size = 9
            .Axes(xlCategory).AxisTitle.Delete
            .Axes(xlValue).AxisTitle.Left = 6
            .Axes(xlValue).AxisTitle.Top = ActiveChart.PlotArea.Height / 3
            .ChartType = xlColumnClustered
        End With
        
        '====================================
        ' ADD UNIT PRICE TABLE + CHART HERE
        '====================================
        
        
    End If
    
    '================================================
    'TO DO:
    '       - Create total qty chart
    '       - Create data + chart for unit price
    '       - Change default chart title to "IMPORTED INTO"
    '   =SUMPRODUCT(A2:A3,B2:B3)/SUM(B2:B3) --> Weighted Average
    '===============================================
    

    '==========================================================='
    Sheets("Main").Activate
    
    'Replace headers to reduce size of printout
        Cells.Replace What:="Value (Dollars)", Replacement:="$ USD", LookAt:= _
            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        Cells.Replace What:="Quantity", Replacement:="QTY", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        Cells.Replace What:="Unit Price", Replacement:="$/Unit", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
    
    'Center Headers
        Rows("5:5").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Range("A5").Select
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
    'Zoom Out
        ActiveWindow.Zoom = 80
        
    'Rename 'All Districts' to Totals
        Range("A6").Select
        ActiveCell.FormulaR1C1 = "Totals"
    
    'Merge HS code & date to larger field
        Range("A1:G1").Select
        With Selection
            .HorizontalAlignment = xlGeneral
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
        End With
        Selection.UnMerge
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Selection.Merge
        Range("A2:G2").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Selection.Merge
        With Selection
            .HorizontalAlignment = xlGeneral
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
        End With
        Selection.UnMerge
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Selection.Merge
        Range("A1:G2").Select
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
        End With
    
    'Autofit Cells
        Cells.Select
        Cells.EntireColumn.AutoFit
    
    'Add printing properties [landscape, narrow margins, etc]
        With ActiveSheet.PageSetup
            .PrintTitleRows = ""
            .PrintTitleColumns = ""
        End With
        ActiveSheet.PageSetup.PrintArea = Range(Cells(1, 1), Cells((LastVertRow - 2), printrange)).Address
            
        With ActiveSheet.PageSetup
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = ""
            .RightFooter = ""
            .LeftMargin = Application.InchesToPoints(0.25)
            .RightMargin = Application.InchesToPoints(0.25)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
            .PrintHeadings = False
            .PrintGridlines = False
            .PrintComments = xlPrintNoComments
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = xlLandscape
            .Draft = False
            .PaperSize = xlPaperLetter
            .FirstPageNumber = xlAutomatic
            .Order = xlDownThenOver
            .BlackAndWhite = False
            .Zoom = 100
            .PrintErrors = xlPrintErrorsDisplayed
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
            .EvenPage.LeftHeader.Text = ""
            .EvenPage.CenterHeader.Text = ""
            .EvenPage.RightHeader.Text = ""
            .EvenPage.LeftFooter.Text = ""
            .EvenPage.CenterFooter.Text = ""
            .EvenPage.RightFooter.Text = ""
            .FirstPage.LeftHeader.Text = ""
            .FirstPage.CenterHeader.Text = ""
            .FirstPage.RightHeader.Text = ""
            .FirstPage.LeftFooter.Text = ""
            .FirstPage.CenterFooter.Text = ""
            .FirstPage.RightFooter.Text = ""
        End With
        
        ActiveWindow.View = xlPageBreakPreview
       ' ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
        ActiveWindow.View = xlNormalView
        
        Cells(1, printrange).Value = Range("B3").Value
        Cells(1, printrange).Font.Bold = True
        
        Sheets("Graph Data").Activate
        ActiveWindow.Zoom = 90
        'Pause for 0.5 seconds. Depends on function defined at top of module
        Application.Wait Now + TimeValue("00:00:01")
        ActiveWindow.Zoom = 100
        Application.Wait Now + TimeValue("00:00:01")
        ActiveWindow.Zoom = 90
        Application.Wait Now + TimeValue("00:00:01")
        ActiveWindow.Zoom = 80
        
        Range("A1").Select

    Else

End If

End Sub
