Sub RandomGraphDijkstra()

'Run time optimisation - disabling of different functions
screenUpdateState = Application.ScreenUpdating
statusBarState = Application.DisplayStatusBar
eventsState = Application.EnableEvents
displayPageBreakState = ActiveSheet.DisplayPageBreaks
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

For q = 15002 To 20001
    
    ActiveWorkbook.Worksheets("MAIN").Cells(1, 22) = 999998
    For n = 1 To 20
        ActiveWorkbook.Worksheets("MAIN").Cells(n, 23) = 0
    Next n
    lastlinework = ActiveWorkbook.Worksheets("WORK").Cells(Rows.Count, 1).End(xlUp).Row
    If lastlinework > 1 Then
        'delete work area
        Sheets("Work").Activate
        ActiveSheet.Range("A2:U" & lastlinework).Select
        Selection.ClearContents
    End If
    m = 0
    ActiveWorkbook.Worksheets("MAIN").Cells(1, 21) = Excel.WorksheetFunction.RandBetween(1, 20)
    ActiveWorkbook.Worksheets("MAIN").Cells(2, 21) = Excel.WorksheetFunction.RandBetween(1, 20)
    For n = 1 To 20
            m = m + 1
            For i = m To 20
                ActiveWorkbook.Worksheets("MAIN").Cells(i, n) = 999999
                ActiveWorkbook.Worksheets("MAIN").Cells(n, i) = 999999
                Count = Count + 1
            Next i
    Next n
    I3 = 4
    For n = 1 To 19
        i2 = Excel.WorksheetFunction.RandBetween(1, I3) 'Random number of edges per Node
        i2 = Excel.WorksheetFunction.Round(i2, 1)
        If i2 <> 0 Then
            For Count = 0 To i2
                i = Excel.WorksheetFunction.RandBetween(n + 1, 20) 'We fill up all cells below ActiveWorkbook.Worksheets("MAIN") diagonal starting at 4
                radno = Excel.WorksheetFunction.RandBetween(1, 99)
                ActiveWorkbook.Worksheets("MAIN").Cells(i, n) = radno
                ActiveWorkbook.Worksheets("MAIN").Cells(n, i) = radno
                Count = Count + 1
            Next Count
        End If
        If n = 5 Then
            I3 = 3
        ElseIf n = 10 Then
            I3 = 2
        ElseIf n = 15 Then
            I3 = 1
        End If
    Next n
    
    'Dijkstra
    abfahrt = ActiveWorkbook.Worksheets("MAIN").Cells(1, 21)
    ende = ActiveWorkbook.Worksheets("MAIN").Cells(2, 21)
    
    If abfahrt = ende Then
        ActiveWorkbook.Worksheets("MAIN").Cells(1, 22) = 0
    Else
        
        ActNode = abfahrt
        Mindistmin = 999998
        ActiveWorkbook.Worksheets("WORK").Cells(2, 1) = 0
        ActiveWorkbook.Worksheets("WORK").Cells(2, 2) = abfahrt
        'While Actualnode is not the same as end node
        While ActNode <> ende
            '20 nodes and therefore 20 entries in the array
            n = 1
            For i = 1 To 20
                If ActiveWorkbook.Worksheets("MAIN").Cells(ActNode, i) < 999999 Then
                    IsAlready = False
                    For x = 2 To 21
                        If ActiveWorkbook.Worksheets("WORK").Cells(2, x) = i Then
                            IsAlready = True
                            Exit For
                        End If
                    Next x
                    If IsAlready = False Then
                        If n = 1 Then
                            'If it is the first connection found on this node then add it to the line
                            lzwischen = ActiveWorkbook.Worksheets("WORK").Cells(2, 1)
                            ActiveWorkbook.Worksheets("WORK").Cells(2, 1) = ActiveWorkbook.Worksheets("WORK").Cells(2, n) + ActiveWorkbook.Worksheets("MAIN").Cells(ActNode, i)
                            For m = 2 To 21
                                If ActiveWorkbook.Worksheets("WORK").Cells(2, m) = "" Then
                                        lastfill = m
                                        Exit For
                                End If
                            Next m
                            ActiveWorkbook.Worksheets("WORK").Cells(2, lastfill) = i
                            n = n + 1
                        Else
                            'If not, save to newly created line
                            'ActiveWorkbook.Worksheets("WORK"): new lastline copy of up to date line and save index
                            INDEC = ActiveWorkbook.Worksheets("WORK").Cells(Rows.Count, 1).End(xlUp).Row + 1
                            For m = 1 To 21
                                If ActiveWorkbook.Worksheets("WORK").Cells(2, m + 1) = "" Then
                                    lastfill = m
                                    Exit For
                                Else
                                    ActiveWorkbook.Worksheets("WORK").Cells(INDEC, m) = ActiveWorkbook.Worksheets("WORK").Cells(2, m)
                                End If
                            Next m
                            ActiveWorkbook.Worksheets("WORK").Cells(INDEC, lastfill) = i
                            ActiveWorkbook.Worksheets("WORK").Cells(INDEC, 1) = lzwischen + ActiveWorkbook.Worksheets("MAIN").Cells(ActNode, i)
                        End If
                    End If
                End If
            Next i
            If n = 1 Then
                    'If no connection is found, we set a high value to the length, so that this path is not at the top anymore
                    ActiveWorkbook.Worksheets("WORK").Cells(2, 1) = 999998
            End If
            'Filter the ActiveWorkbook.Worksheets("WORK") sheet ascending by first row (length)
            ActiveWorkbook.Worksheets("WORK").AutoFilter.Sort.SortFields.Clear
            ActiveWorkbook.Worksheets("WORK").AutoFilter.Sort.SortFields.Add Key:=Range( _
                "A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
                xlSortNormal
            With ActiveWorkbook.Worksheets("WORK").AutoFilter.Sort
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            'Find last node inserted in shortest path (top row now)
            If ActiveWorkbook.Worksheets("WORK").Cells(2, 1) >= 999998 Then
                Exit Sub
            End If
            For i = 2 To 21
                If ActiveWorkbook.Worksheets("WORK").Cells(2, i) = "" Then
                    ActNode = ActiveWorkbook.Worksheets("WORK").Cells(2, (i - 1))
                    Exit For
                End If
            Next i
        Wend
        ActiveWorkbook.Worksheets("MAIN").Cells(1, 22) = ActiveWorkbook.Worksheets("WORK").Cells(2, 1)
        For n = 1 To 20
            If ActiveWorkbook.Worksheets("WORK").Cells(2, n + 1) <> "" Then
                ActiveWorkbook.Worksheets("MAIN").Cells(n, 23) = ActiveWorkbook.Worksheets("WORK").Cells(2, n + 1)
            Else
                Exit For
            End If
        Next n
    End If
    
    g = 1
    
    For n = 1 To 20
        For m = 1 To 20
            ActiveWorkbook.Worksheets("OUTPUT").Cells(q, g) = ActiveWorkbook.Worksheets("MAIN").Cells(m, n)
            g = g + 1
        Next m
    Next n
    ActiveWorkbook.Worksheets("OUTPUT").Cells(q, g) = ActiveWorkbook.Worksheets("MAIN").Cells(1, 21)
    g = g + 1
    ActiveWorkbook.Worksheets("OUTPUT").Cells(q, g) = ActiveWorkbook.Worksheets("MAIN").Cells(2, 21)
    g = g + 1
    ActiveWorkbook.Worksheets("OUTPUT").Cells(q, g) = ActiveWorkbook.Worksheets("MAIN").Cells(1, 22)
    g = g + 1
    For m = 1 To 20
            ActiveWorkbook.Worksheets("OUTPUT").Cells(q, g) = ActiveWorkbook.Worksheets("MAIN").Cells(m, 23)
            g = g + 1
    Next m
Next q

'enabling the features disabled for the runtime optimization
Application.ScreenUpdating = screenUpdateState
Application.DisplayStatusBar = statusBarState
Application.EnableEvents = eventsState
ActiveSheet.DisplayPageBreaks = displayPageBreaksState

End Sub
