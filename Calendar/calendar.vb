Sub MakeCalendar()
  ' === Vars ===
  var_month = 1
  var_year = 2026

  ' Print 2 pages, 6 months on each one
  TwoPage = False
  ' Print months in 2 columns
  CreateMonthsSheet = False
  ShowWeeks = False
  PrintMonthPerPage = False

  HexValues = False
  ShowGrid = True
  AdjustPrint = True

  If CreateMonthsSheet Then
    TwoPage = False
  End If

  weekColOff = 0
  If ShowWeeks Then
    weekColOff = 1
  End If

  startRow = 2
  startCol = 2

  rowsInMonth = 8
  colsInMonth = 7 + weekColOff

  monthRows = 3
  monthCols = 4
  If TwoPage Then
    monthRows = 4
    monthCols = 3
  End If
  If CreateMonthsSheet Then
    monthRows = 6
    monthCols = 2
  End If

  ' -------------------
  monr = startRow
  monc = startCol

  dayr = startRow + 2
  dayc = startCol + weekColOff

  firstRow = startRow
  firstCol = startCol

  lastRow = firstRow + rowsInMonth * monthRows - 1
  lastCol = firstCol + colsInMonth * monthCols - 1

  Dim weekdays(7) As String
  weekdays(1) = "Mon"
  weekdays(2) = "Tue"
  weekdays(3) = "Wed"
  weekdays(4) = "Thu"
  weekdays(5) = "Fri"
  weekdays(6) = "Sat"
  weekdays(7) = "Sun"


  ' ----------------------------------------------
  Sheets.Add Before:=ActiveSheet
  Application.ScreenUpdating = False
  Cells.ClearContents
  Cells.ClearFormats
  ActiveWindow.View = xlNormalView

  SheetName = CStr(var_year) & "-Full"
  If TwoPage Then
      SheetName = CStr(var_year) & "-TwoPage"
  End If
  If CreateMonthsSheet Then
      SheetName = CStr(var_year) & "-BigMonths"
  End If
  If HexValues Then
    SheetName = SheetName & "Hex"
  End If

  sheet_exists = False
  For i = 1 To Worksheets.Count
    If Worksheets(i).Name = SheetName Then
      sheet_exists = True
    End If
  Next i
  If Not sheet_exists Then
    ActiveSheet.Name = SheetName
  End If

  WeekNum = 0

  For i = 1 To monthRows
    For j = 1 To monthCols
      DateString = CStr(var_month) & "-" & CStr(var_year)
      StartDay = DateValue(DateString)

      Set var_monthCell = Range(Cells(monr, monc), Cells(monr, monc))

      ' Month name
      With Range(Cells(monr, monc), Cells(monr, monc + colsInMonth - 1))
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Font.Size = 18
        .Font.Bold = True
        .RowHeight = 25
      End With

      ' Print english months
      var_monthCell.Value = Application.Text(StartDay, "[$-409]mmmm")

      ' Day names
      With Range(Cells(monr + 1, monc), Cells(monr + 1, monc + colsInMonth - 1))
        .ColumnWidth = 5
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = xlHorizontal
        .Font.Size = 14
        .Font.Bold = True
        .RowHeight = 20
        .Interior.ColorIndex = 15
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
      End With

      If HexValues Then
        Range(Cells(monr + 1, monc), Cells(monr + 1, monc + colsInMonth - 1)).ColumnWidth = 5.5
      End If

      If ShowWeeks Then
        Range(Cells(monr + 1, monc), Cells(monr + 1, monc)).ColumnWidth = 3
      End If

      Range(Cells(monr + 1, monc + weekColOff + 0), Cells(monr + 1, monc + weekColOff + 0)) = weekdays(1)
      Range(Cells(monr + 1, monc + weekColOff + 1), Cells(monr + 1, monc + weekColOff + 1)) = weekdays(2)
      Range(Cells(monr + 1, monc + weekColOff + 2), Cells(monr + 1, monc + weekColOff + 2)) = weekdays(3)
      Range(Cells(monr + 1, monc + weekColOff + 3), Cells(monr + 1, monc + weekColOff + 3)) = weekdays(4)
      Range(Cells(monr + 1, monc + weekColOff + 4), Cells(monr + 1, monc + weekColOff + 4)) = weekdays(5)
      Range(Cells(monr + 1, monc + weekColOff + 5), Cells(monr + 1, monc + weekColOff + 5)) = weekdays(6)
      Range(Cells(monr + 1, monc + weekColOff + 6), Cells(monr + 1, monc + weekColOff + 6)) = weekdays(7)

      ' Day numbers
      With Range(Cells(dayr, dayc), Cells(dayr + rowsInMonth - 3, dayc + colsInMonth - weekColOff))
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .Font.Size = 14
        .RowHeight = 25
        .NumberFormat = "@"
      End With

      If HexValues Then
        Range(Cells(dayr, dayc), Cells(dayr + rowsInMonth - 3, dayc + colsInMonth - weekColOff)).HorizontalAlignment = xlCenter
      End If

      ' Week numbers
      If ShowWeeks Then
        With Range(Cells(dayr, dayc - 1), Cells(dayr + rowsInMonth - 3, dayc - 1))
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .Font.Size = 10
          .Font.Bold = True
        End With
      End If

      DayOfWeek = Weekday(StartDay, 2)
      CurYear = Year(StartDay)
      CurMonth = Month(StartDay)
      FinalDay = DateSerial(CurYear, CurMonth + 1, 1)

      dayOneOffset = DayOfWeek - 1
      Cells(dayr, dayc + dayOneOffset).Value = 1

      CurDay = 0

      For Each cell In Range(Cells(dayr, dayc), Cells(dayr + 5, dayc + colsInMonth - weekColOff - 1))
        CellRow = cell.Row
        CellCol = cell.Column
        PrintValues = True

        If CellRow = dayr And CellCol < dayc + dayOneOffset Then
          PrintValues = False
        End If

        If PrintValues = True Then
          CurDay = CurDay + 1
          cell.Value = CurDay
          If HexValues Then
            hex_ = ""
            If CurDay < 16 Then
              hex_ = "0"
            End If
            cell.Value = hex_ & Hex(CurDay)
          End If

          If CurDay > (FinalDay - StartDay) Then
            cell.Value = ""
            Exit For
          End If

          cell.Borders(xlEdgeLeft).LineStyle = xlContinuous
          cell.Borders(xlEdgeLeft).Weight = xlHairline
          cell.Borders(xlEdgeRight).LineStyle = xlContinuous
          cell.Borders(xlEdgeRight).Weight = xlHairline
          cell.Borders(xlEdgeTop).LineStyle = xlContinuous
          cell.Borders(xlEdgeTop).Weight = xlHairline
          cell.Borders(xlEdgeBottom).LineStyle = xlContinuous
          cell.Borders(xlEdgeBottom).Weight = xlHairline
        End If

        If ShowWeeks Then
          If CellRow >= dayr And CellCol = dayc Then
            If PrintValues Then
              WeekNum = WeekNum + 1
            End If
            Cells(cell.Row, dayc - 1).Value = WeekNum
          End If
        End If
      Next

      ' --------
      monc = monc + colsInMonth
      dayc = dayc + colsInMonth
      var_month = var_month + 1
    Next j

    monr = monr + rowsInMonth
    monc = startCol
    dayr = dayr + rowsInMonth
    dayc = startCol + weekColOff
  Next i


  '************************************************** Grid **********************************
  If ShowGrid Then
    Set var_wholeRange = Range(Cells(firstRow, firstCol), Cells(lastRow, lastCol))

    var_wholeRange.Borders(xlEdgeLeft).LineStyle = xlContinuous
    var_wholeRange.Borders(xlEdgeLeft).Weight = xlThin
    var_wholeRange.Borders(xlEdgeTop).LineStyle = xlContinuous
    var_wholeRange.Borders(xlEdgeTop).Weight = xlThin
    var_wholeRange.Borders(xlEdgeRight).LineStyle = xlContinuous
    var_wholeRange.Borders(xlEdgeRight).Weight = xlThin
    var_wholeRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
    var_wholeRange.Borders(xlEdgeBottom).Weight = xlThin

    For i = 1 To monthCols - 1
      row1 = firstRow
      row2 = lastRow
      borderCol = firstCol + i * colsInMonth - 1

      Set var_columnRange_1 = Range(Cells(row1, borderCol), Cells(row2, borderCol))
      Set var_columnRange_2 = Range(Cells(row1, borderCol + 1), Cells(row2, borderCol + 1))

      var_columnRange_1.Borders(xlEdgeRight).LineStyle = xlContinuous
      var_columnRange_1.Borders(xlEdgeRight).Weight = xlThin
      var_columnRange_2.Borders(xlEdgeLeft).LineStyle = xlContinuous
      var_columnRange_2.Borders(xlEdgeLeft).Weight = xlThin
    Next i

    For i = 1 To monthRows - 1
      col1 = firstCol
      col2 = lastCol
      borderRow = firstRow + i * rowsInMonth - 1

      Set var_rowRange_1 = Range(Cells(borderRow, col1), Cells(borderRow, col2))
      Set var_rowRange_2 = Range(Cells(borderRow + 1, col1), Cells(borderRow + 1, col2))

      var_rowRange_1.Borders(xlEdgeBottom).LineStyle = xlContinuous
      var_rowRange_1.Borders(xlEdgeBottom).Weight = xlThin
      var_rowRange_2.Borders(xlEdgeTop).LineStyle = xlContinuous
      var_rowRange_2.Borders(xlEdgeTop).Weight = xlThin
    Next i
  End If

  '*** Year Cell
  Cells(1, 1).Value = var_year
  Cells(1, 1).ColumnWidth = 12
  Cells(1, 1).HorizontalAlignment = xlCenterAcrossSelection
  Cells(1, 1).VerticalAlignment = xlCenter

  Cells(1, 1).Font.Size = 22
  Cells(1, 1).Font.Bold = True
  Cells(1, 1).Font.Color = &HFFFFFF
  Cells(1, 1).Font.TintAndShade = 0


  '************************************************** Print Properties **********************************
  If AdjustPrint Then
    ActiveWindow.View = xlPageBreakPreview
    ActiveWindow.Zoom = 85

    ActiveSheet.PageSetup.PrintArea = Range(Cells(firstRow, firstCol), Cells(lastRow, lastCol)).Address
    With ActiveSheet.PageSetup
      .LeftMargin = Application.CentimetersToPoints(0.5)
      .RightMargin = Application.CentimetersToPoints(0.5)
      .TopMargin = 0
      .BottomMargin = 0
      .HeaderMargin = 0
      .FooterMargin = 0
      .CenterHorizontally = True
      .CenterVertically = True
      .Orientation = xlLandscape
      .Zoom = False
      .FitToPagesWide = 1
      .FitToPagesTall = 1
      .PaperSize = xlPaperA4
      .Order = xlOverThenDown

      ' .Draft = False
      ' .FirstPageNumber = xlAutomatic
      ' .BlackAndWhite = False
      ' .PrintErrors = xlPrintErrorsDisplayed
      ' .PrintHeadings = False
      ' .PrintGridlines = False
      ' .PrintComments = xlPrintNoComments
      ' .PrintQuality = 600
    End With

    If PrintMonthPerPage Then
      ActiveSheet.ResetAllPageBreaks
      For i = 1 To monthRows - 1
          breakRow = firstRow + i * rowsInMonth
          ActiveSheet.HPageBreaks.Add Before:=Cells(breakRow, 1)
      Next i
      For j = 1 To monthCols - 1
          breakCol = firstCol + j * colsInMonth
          ActiveSheet.VPageBreaks.Add Before:=Cells(1, breakCol)
      Next j

      ActiveSheet.PageSetup.Zoom = 250
    End If
  End If

  If TwoPage Then
    ActiveSheet.PageSetup.Zoom = 100
    ActiveSheet.HPageBreaks.Add Cells(firstRow + 2 * rowsInMonth, 1)
  End If

  If CreateMonthsSheet And Not HexValues Then
    With ActiveSheet.PageSetup
      .Zoom = 130
      .CenterHorizontally = True
      .CenterVertically = False
      .TopMargin = 0
      .BottomMargin = 0
      .LeftMargin = 0
      .RightMargin = 0
      .HeaderMargin = 0
      .FooterMargin = 0
    End With

    ActiveSheet.HPageBreaks.Add Cells(firstRow + 2 * rowsInMonth, 1)
    ActiveSheet.HPageBreaks.Add Cells(firstRow + 4 * rowsInMonth, 1)
  End If
End Sub
