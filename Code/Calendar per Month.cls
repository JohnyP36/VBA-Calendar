VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Blad4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CalendarMaker()

       ' Unprotect sheet if had previous calendar to prevent error.
       ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False
       ' Prevent screen flashing while drawing calendar.
       Application.ScreenUpdating = False
       ' Set up error trapping.
       On Error GoTo MyErrorTrap
       ' Clear area a1:g14 including any previous calendar.
       Range("a1:l21").Delete
       ' Use InputBox to get desired month and year and set variable MyInput.
       MyInput = InputBox("Geeft de maand en het jaar voor de kalender op. . . ")
       ' Allow user to end macro with Cancel in InputBox.
       If MyInput = "" Then Exit Sub
       ' Get the date value of the beginning of inputted month.
       StartDay = DateValue(MyInput)
       ' Check if valid date but not the first of the month
       ' -- if so, reset StartDay to first day of month.
       If Day(StartDay) <> 1 Then
           StartDay = DateValue(Month(StartDay) & "/1/" & _
               Year(StartDay))
       End If
       ' Prepare cell for Month and Year as fully spelled out.
       Range("a1").NumberFormat = "mmmm yyyy"
       ' Center the Month and Year label across a1:g1 with appropriate size, height and bolding.
       With Range("a1:g1")
           .HorizontalAlignment = xlCenterAcrossSelection
           .VerticalAlignment = xlCenter
           .Font.Size = 18
           .Font.Bold = True
           .RowHeight = 35
       End With
       ' Change properties of cell where Month and Year are spelled out
       Range("A1").Select
       With Selection.Font
           .Name = "MV Boli"
           .Size = 21
           .Strikethrough = False
           .Superscript = False
           .Subscript = False
           .OutlineFont = False
           .Shadow = False
           .Underline = xlUnderlineStyleSingle
           .Color = -16777012
           .TintAndShade = 0
           .ThemeFont = xlThemeFontNone
       End With
       ' Prepare a2:g2 for day of week labels with centering, size, height and bolding.
       With Range("a2:g2")
           .ColumnWidth = 20
           .VerticalAlignment = xlCenter
           .HorizontalAlignment = xlCenter
           .VerticalAlignment = xlCenter
           .Orientation = xlHorizontal
           .Font.Size = 12
           .Font.Bold = True
           .RowHeight = 20
       End With
       ' Put days of week in a2:g2.
       Range("a2") = "Zondag"
       Range("b2") = "Maandag"
       Range("c2") = "Dinsdag"
       Range("d2") = "Woensdag"
       Range("e2") = "Donderdag"
       Range("f2") = "Vrijdag"
       Range("g2") = "Zaterdag"
       ' Prepare a3:g7 for dates with left/top alignment, size, height and bolding.
       With Range("a3:g8")
           .HorizontalAlignment = xlRight
           .VerticalAlignment = xlTop
           .Font.Size = 18
           .Font.Bold = True
           .RowHeight = 21
           .Font.ColorIndex = 11
       End With
       ' Put inputted month and year fully spelling out into "a1".
       Range("a1").Value = Application.Text(MyInput, "mmmm yyyy")
       ' Set variable and get which day of the week the month starts.
       DayofWeek = Weekday(StartDay)
       ' Set variables to identify the year and month as separate variables.
       CurYear = Year(StartDay)
       CurMonth = Month(StartDay)
       ' Set variable and calculate the first day of the next month.
       FinalDay = DateSerial(CurYear, CurMonth + 1, 1)
       ' Place a "1" in cell position of the first day of the chosen month based on DayofWeek.
       Select Case DayofWeek
           Case 1
               Range("a3").Value = 1
           Case 2
               Range("b3").Value = 1
           Case 3
               Range("c3").Value = 1
           Case 4
               Range("d3").Value = 1
           Case 5
               Range("e3").Value = 1
           Case 6
               Range("f3").Value = 1
           Case 7
               Range("g3").Value = 1
       End Select
       ' Loop through range a3:g8 incrementing each cell after the "1" cell.
       For Each Cell In Range("a3:g8")
           RowCell = Cell.Row
           ColCell = Cell.Column
           ' Do if "1" is in first column.
           If Cell.Column = 1 And Cell.Row = 3 Then
           ' Do if current cell is not in 1st column.
           ElseIf Cell.Column <> 1 Then
               If Cell.Offset(0, -1).Value >= 1 Then
                   Cell.Value = Cell.Offset(0, -1).Value + 1
                   ' Stop when the last day of the month has been entered.
                   If Cell.Value > (FinalDay - StartDay) Then
                       Cell.Value = ""
                       ' Exit loop when calendar has correct number of days shown.
                       Exit For
                   End If
               End If
           ' Do only if current cell is not in Row 3 and is in Column 1.
           ElseIf Cell.Row > 3 And Cell.Column = 1 Then
               Cell.Value = Cell.Offset(-1, 6).Value + 1
               ' Stop when the last day of the month has been entered.
               If Cell.Value > (FinalDay - StartDay) Then
                   Cell.Value = ""
                   ' Exit loop when calendar has correct number of days shown.
                   Exit For
               End If
           End If
       Next

       ' Create Entry cells, format them centered, wrap text, and border around days.
       For x = 0 To 5
           Range("A4").Offset(x * 2, 0).EntireRow.Insert
           With Range("A4:G4").Offset(x * 2, 0)
               .RowHeight = 65
               .HorizontalAlignment = xlCenter
               .VerticalAlignment = xlTop
               .WrapText = True
               .Font.Size = 10
               .Font.Bold = False
               ' Unlock these cells to be able to enter text later after sheet is protected.
               .Locked = False
           End With
           ' Put border around the block of dates.
           With Range("A3").Offset(x * 2, 0).Resize(2, _
           7).Borders(xlLeft)
               .Weight = xlThick
               .ColorIndex = xlAutomatic
           End With

           With Range("A3").Offset(x * 2, 0).Resize(2, _
           7).Borders(xlRight)
               .Weight = xlThick
               .ColorIndex = xlAutomatic
           End With
           Range("A3").Offset(x * 2, 0).Resize(2, 7).BorderAround _
              Weight:=xlThick, ColorIndex:=xlAutomatic
       Next
       
       For y = 0 To 6
           Range("A3:A14").Select
           With Range("A3:A14").Interior
               .ColorIndex = 15
           End With
       Next
       For y = 0 To 6
           Range("G3:G14").Select
           With Range("G3:G14").Interior
               .ColorIndex = 15
           End With
       Next
       
       Call VoegWeeknummersToe
       Call SplitDagenInTweeën
       
       If Range("B18").Value = "" Then Range("B18").Offset(0, 0) _
          .Resize(3, 8).EntireRow.Delete
       ' Turn off gridlines.
       ActiveWindow.DisplayGridlines = False
       ' Protect sheet to prevent overwriting the dates.
       ActiveSheet.Protect DrawingObjects:=True, Contents:=True, _
          Scenarios:=True
       ' Resize window to show all of calendar (may have to be adjusted
       ' for video configuration).
       ActiveWindow.WindowState = xlMaximized
       ActiveWindow.ScrollRow = 1
       ' Allow screen to redraw with calendar showing.
       Application.ScreenUpdating = True
       ' Prevent going to error trap unless error found by exiting Sub here.
       Exit Sub
   ' Error causes msgbox to indicate the problem, provides new input box,
   ' and resumes at the line that caused the error.
MyErrorTrap:
       MsgBox "Misschien heb u de maand en het jaar niet correct ingevoerd." _
           & Chr(13) & "Spel de maand juist" _
           & " (of gebruik de 3-letterige afkorting)" _
           & Chr(13) & "en type het jaar met 4 getallen"
       MyInput = InputBox("Geeft de maand en het jaar op voor de kalender")
       If MyInput = "" Then Exit Sub
       Resume
End Sub

Sub VoegWeeknummersToe()
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Selection.ColumnWidth = 4.64
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
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "=ISOWEEKNUM(RC[10])"
    Range("A5").Select
    ActiveCell.FormulaR1C1 = "=ISOWEEKNUM(RC[10])"
    Range("A7").Select
    ActiveCell.FormulaR1C1 = "=ISOWEEKNUM(RC[10])"
    Range("A9").Select
    ActiveCell.FormulaR1C1 = "=ISOWEEKNUM(RC[10])"
    Range("A11").Select
    ActiveCell.FormulaR1C1 = "=ISOWEEKNUM(RC[10]+1)"
    Range("A13").Select
    ActiveCell.FormulaR1C1 = "=ISOWEEKNUM(RC[10]+1)"
    Range("A13,A11,A9,A7,A5,A3").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A4,A6,A8,A10,A12,A14").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A2").Select
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
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveCell.FormulaR1C1 = "week"
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "=CONCAT(RC[-9],"" "",R3C12)"
    Range("K5").Select
    ActiveCell.FormulaR1C1 = "=CONCAT(RC[-9],"" "",R3C12)"
    Range("K7").Select
    ActiveCell.FormulaR1C1 = "=CONCAT(R[-2]C[-9],"" "",R3C12)"
    Range("K9").Select
    ActiveCell.FormulaR1C1 = "=CONCAT(RC[-9],"" "",R3C12)"
    Range("K11").Select
    ActiveCell.FormulaR1C1 = "=CONCAT(RC[-9],"" "",R3C12)"
    Range("K13").Select
    ActiveCell.FormulaR1C1 = "=CONCAT(RC[-9],"" "",R3C12)"
    Range("L3").Select
    ActiveCell.FormulaR1C1 = "=TEXT(R[-2]C[-10],""mmmm jjjj"")"
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "=CONCAT(RC[-3],"" "",R3C12)"
    Range("K5").Select
    ActiveCell.FormulaR1C1 = "=CONCAT(RC[-3],"" "",R3C12)"
    Range("K7").Select
    ActiveCell.FormulaR1C1 = "=CONCAT(RC[-3],"" "",R3C12)"
    Range("K9").Select
    ActiveCell.FormulaR1C1 = "=CONCAT(RC[-3],"" "",R3C12)"
    Columns("K:L").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Range("A3,A5,A7,A9,A11,A13").Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .Color = 6737151
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub SplitDagenInTweeën()
' Split cells under each day in half (add an extra row)
    Rows("4:4").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    Rows("7:7").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    Rows("10:10").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    ActiveWindow.SmallScroll Down:=6
    Rows("13:13").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    Rows("16:16").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    ActiveWindow.SmallScroll Down:=3
    Rows("19:19").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    Range("19:20,16:17,13:14,10:11").Select
    ActiveWindow.SmallScroll Down:=-9
    Range("19:20,16:17,13:14,10:11,7:8,4:5").Select
    Selection.RowHeight = 32.5
    Range("B4:H5,B7:H8,B10:H11,B13:H14,B16:H17,B19:H20").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlDash
        .ThemeColor = 1
        .TintAndShade = -0.14996795556505
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDash
        .ThemeColor = 1
        .TintAndShade = -0.14996795556505
        .Weight = xlThin
    End With
    
    'Maakt tekstkleur zwart
    Range("B4:H5,B7:H8,B10:H11,B13:H14,B16:H17,B19:H20").Select
    With Selection.Font
        .ColorIndex = 1
        .TintAndShade = 0
    End With
    Range("A1").Select
End Sub
