# zbrplan-macro
Format zbrplan output for planning
Sub zbrplan()
'
' zbrplan Macro
'
runmacrozbrplan = MsgBox("Run ZBRPLAN formatting program?", vbYesNo)

If runmacrozbrplan = vbYes Then
    Range("A1:AP2321").Select
    ActiveSheet.Range("$A$1:$AP$2000").RemoveDuplicates Columns:=7, Header:= _
        xlYes
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    Columns("AG:AG").Select
    Selection.Cut
    ActiveWindow.LargeScroll ToRight:=-1
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight
    Columns("P:P").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "net stk"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "ss ok"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "IF(P2>=MAX(S2:AB2), ""OK"")"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[15]>=MAX(RC[18]:RC[27]), ""OK"")"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-1]+RC[-6]+RC[-5]"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]-RC[-4]"
    Range("Q3").Select
    ActiveWindow.Zoom = 90
    ActiveWindow.Zoom = 80
    Cells.Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Font
        .Name = "Calibri"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Range("D3").Select
    Rows("1:1").RowHeight = 49.5
    Columns("S:AB").Select
    Selection.ColumnWidth = 4.43
    Columns("P:P").Select
    Selection.NumberFormat = "0_);[Red](0)"
    Columns("Q:Q").Select
    Selection.NumberFormat = "0.00"
    Selection.NumberFormat = "0_);[Red](0)"
    Columns("F:F").ColumnWidth = 30.57
    Range("Q2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 14083324
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("P2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15917529
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("I2").Select
    Cells.FormatConditions.Delete
    Range("I2").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("U2:AB2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15921906
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("U7").Select
   lastrow = Cells(Rows.Count, 2).End(xlUp).Row
    Range("A2").Select
    Selection.AutoFill Destination:=Range("A2:A" & lastrow)
    Range("P2:Q2").Select
    Selection.AutoFill Destination:=Range("P2:Q" & lastrow)
     Range("U2:AB2").Select
    Selection.AutoFill Destination:=Range("U2:AB" & lastrow), Type:=xlFillFormats
     Columns("AO:AS").Select
    Selection.Cut
    Columns("AC:AC").Select
    Selection.Insert Shift:=xlToRight
    Columns("AC:AG").Select
    Selection.ColumnWidth = 3.3
    Columns("AQ:AQ").Select
    Selection.Cut
    Columns("AJ:AJ").Select
    Selection.Insert Shift:=xlToRight
    Range("AM2").Select
    If Range("an2") = 7010 Then
        Range("Am2") = "=VLOOKUP(H2,'H:\DRP\[Branch Planning Notes.xlsx]Australia'!$A$3:$B$255,2,0)"
        Else: If Range("an2") = 7050 Then Range("AM2") = "=VLOOKUP(H2,'H:\DRP\[Branch Planning Notes.xlsx]Singapore'!$A$3:$B$255,2,0)"

    End If
      Range("AM2").Select
    Selection.AutoFill Destination:=Range("AM2:AM" & lastrow)
    Range("AM2:AM" & lastrow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Range("L1").Select
    Selection.Style = "Neutral"
    Worksheets("Sheet1").Columns("AM").AutoFit
     With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    Columns("AI").ColumnWidth = 4.29
     Columns("AK").ColumnWidth = 4.86
     Columns("I:K").ColumnWidth = 6.29
    ActiveWindow.FreezePanes = True
    Selection.AutoFilter
    Range("l2").Select
Else
End If

End Sub
