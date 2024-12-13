Attribute VB_Name = "Module1"

Sub lengthTime()
Dim zeroFlag As Integer
Dim lengthSize As Long
zeroFlag = 0
lengthSize = ActiveSheet.Range("B7").Value * ActiveSheet.Range("B9").Value
If Sheets("Calculate").Range("B6").Value = "annually" Then
        lengthSize = 1 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "semiannually" Then
        lengthSize = 2 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "quarterly" Then
        lengthSize = 4 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "monthly" Then
       lengthSize = 12 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "daily" Then
        lengthSize = 365 * ActiveSheet.Range("B9").Value
    End If
Range("C11").Value = zeroFlag
For i = 1 To lengthSize
    ActiveSheet.Cells(11, 3 + i).Value = i
Next i
End Sub
Sub payment()
Dim lengthSize As Long
Dim DAType As String
Dim increFactor As Double
Dim decreFactor As Double
Dim growthFactor As Double
Dim degrowthFactor As Double
Dim deferredStart As Integer
deferredStart = ActiveSheet.Range("D3").Value
If Sheets("Calculate").Range("B6").Value = "annually" Then
        lengthSize = 1 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "semiannually" Then
        lengthSize = 2 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "quarterly" Then
        lengthSize = 4 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "monthly" Then
       lengthSize = 12 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "daily" Then
        lengthSize = 365 * ActiveSheet.Range("B9").Value
    End If
DAType = ActiveSheet.Cells(3, 2).Value
If ActiveSheet.Range("C3").Value = "n" Then
    If ActiveSheet.Range("C5").Value = 0 And ActiveSheet.Range("D5").Value = 0 And ActiveSheet.Range("E5").Value = 0 And ActiveSheet.Range("F5").Value = 0 Then
        If DAType = "advance" Then
            For i = 1 To lengthSize
                ActiveSheet.Cells(12, 2 + i).Value = ActiveSheet.Cells(5, 2).Value
            Next i
        ElseIf DAType = "arrear" Then
            For i = 1 To lengthSize
            ActiveSheet.Cells(12, 3 + i).Value = ActiveSheet.Cells(5, 2).Value
            Next i
        End If
    Else
        If ActiveSheet.Range("C5").Value <> 0 And ActiveSheet.Range("D5").Value = 0 And ActiveSheet.Range("E5").Value = 0 And ActiveSheet.Range("F5").Value = 0 Then
            If DAType = "advance" Then
                ActiveSheet.Range("C12").Value = ActiveSheet.Range("B5").Value
                For i = 1 To (lengthSize - 1)
                increFactor = ActiveSheet.Range("C5").Value * i
                ActiveSheet.Cells(12, 3 + i).Value = ActiveSheet.Range("B5").Value + increFactor
                Next i
            ElseIf DAType = "arrear" Then
                ActiveSheet.Range("D12").Value = ActiveSheet.Range("B5").Value
                For i = 1 To (lengthSize - 1)
                increFactor = ActiveSheet.Range("C5").Value * i
                ActiveSheet.Cells(12, 4 + i).Value = ActiveSheet.Range("B5").Value + increFactor
                Next i
            End If
        ElseIf ActiveSheet.Range("D5").Value <> 0 And ActiveSheet.Range("C5").Value = 0 And ActiveSheet.Range("E5").Value = 0 And ActiveSheet.Range("F5").Value = 0 Then
            If DAType = "advance" Then
                ActiveSheet.Range("C12").Value = ActiveSheet.Range("B5").Value
                For i = 1 To (lengthSize - 1)
                decreFactor = ActiveSheet.Range("D5").Value * i
                ActiveSheet.Cells(12, 3 + i).Value = ActiveSheet.Range("B5").Value - decreFactor
                Next i
            ElseIf DAType = "arrear" Then
                ActiveSheet.Range("D12").Value = ActiveSheet.Range("B5").Value
                For i = 1 To (lengthSize - 1)
                decreFactor = ActiveSheet.Range("D5").Value * i
                ActiveSheet.Cells(12, 4 + i).Value = ActiveSheet.Range("B5").Value - decreFactor
                Next i
            End If
        ElseIf ActiveSheet.Range("D5").Value = 0 And ActiveSheet.Range("C5").Value = 0 And ActiveSheet.Range("E5").Value <> 0 And ActiveSheet.Range("F5").Value = 0 Then
            If DAType = "advance" Then
                ActiveSheet.Range("C12").Value = ActiveSheet.Range("B5").Value
                For i = 1 To (lengthSize - 1)
                growthFactor = (1 + ActiveSheet.Range("E5").Value) ^ i
                ActiveSheet.Cells(12, 3 + i).Value = ActiveSheet.Range("B5").Value * growthFactor
                Next i
            ElseIf DAType = "arrear" Then
                ActiveSheet.Range("D12").Value = ActiveSheet.Range("B5").Value
                For i = 1 To (lengthSize - 1)
                growthFactor = (1 + ActiveSheet.Range("E5").Value) ^ i
                ActiveSheet.Cells(12, 4 + i).Value = ActiveSheet.Range("B5").Value * growthFactor
                Next i
            End If
        Else
            If DAType = "advance" Then
                ActiveSheet.Range("C12").Value = ActiveSheet.Range("B5").Value
                For i = 1 To (lengthSize - 1)
                degrowthFactor = (1 - ActiveSheet.Range("F5").Value) ^ i
                ActiveSheet.Cells(12, 3 + i).Value = ActiveSheet.Range("B5").Value * degrowthFactor
                Next i
            ElseIf DAType = "arrear" Then
                ActiveSheet.Range("D12").Value = ActiveSheet.Range("B5").Value
                For i = 1 To (lengthSize - 1)
                degrowthFactor = (1 - ActiveSheet.Range("F5").Value) ^ i
                ActiveSheet.Cells(12, 4 + i).Value = ActiveSheet.Range("B5").Value * degrowthFactor
                Next i
            End If
        End If
    End If
ElseIf ActiveSheet.Range("C3") = "y" Then
    If ActiveSheet.Range("C5").Value = 0 And ActiveSheet.Range("D5").Value = 0 And ActiveSheet.Range("E5").Value = 0 And ActiveSheet.Range("F5").Value = 0 Then
        If DAType = "advance" Then
            For i = 1 To (lengthSize - deferredStart)
                ActiveSheet.Cells(12, 2 + deferredStart + i).Value = ActiveSheet.Cells(5, 2).Value
            Next i
        ElseIf DAType = "arrear" Then
            For i = 1 To (lengthSize - deferredStart)
            ActiveSheet.Cells(12, 3 + deferredStart + i).Value = ActiveSheet.Cells(5, 2).Value
            Next i
        End If
    Else
        If ActiveSheet.Range("C5").Value <> 0 And ActiveSheet.Range("D5").Value = 0 And ActiveSheet.Range("E5").Value = 0 And ActiveSheet.Range("F5").Value = 0 Then
            If DAType = "advance" Then
                ActiveSheet.Cells(12, 3 + deferredStart).Value = ActiveSheet.Range("B5").Value
                For i = 1 To (lengthSize - deferredStart - 1)
                increFactor = ActiveSheet.Range("C5").Value * i
                ActiveSheet.Cells(12, 3 + deferredStart + i).Value = ActiveSheet.Range("B5").Value + increFactor
                Next i
            ElseIf DAType = "arrear" Then
                ActiveSheet.Cells(12, 4 + deferredStart).Value = ActiveSheet.Range("B5").Value
                For i = 1 To (lengthSize - deferredStart - 1)
                increFactor = ActiveSheet.Range("C5").Value * i
                ActiveSheet.Cells(12, 4 + deferredStart + i).Value = ActiveSheet.Range("B5").Value + increFactor
                Next i
            End If
        ElseIf ActiveSheet.Range("D5").Value <> 0 And ActiveSheet.Range("C5").Value = 0 And ActiveSheet.Range("E5").Value = 0 And ActiveSheet.Range("F5").Value = 0 Then
            If DAType = "advance" Then
                ActiveSheet.Cells(12, 3 + deferredStart).Value = ActiveSheet.Range("B5").Value
                For i = 1 To (lengthSize - deferredStart - 1)
                decreFactor = ActiveSheet.Range("D5").Value * i
                ActiveSheet.Cells(12, 3 + deferredStart + i).Value = ActiveSheet.Range("B5").Value - decreFactor
                Next i
            ElseIf DAType = "arrear" Then
                ActiveSheet.Cells(12, 4 + deferredStart).Value = ActiveSheet.Range("B5").Value
                For i = 1 To (lengthSize - deferredStart - 1)
                decreFactor = ActiveSheet.Range("D5").Value * i
                ActiveSheet.Cells(12, 4 + deferredStart + i).Value = ActiveSheet.Range("B5").Value - decreFactor
                Next i
            End If
        ElseIf ActiveSheet.Range("D5").Value = 0 And ActiveSheet.Range("C5").Value = 0 And ActiveSheet.Range("E5").Value <> 0 And ActiveSheet.Range("F5").Value = 0 Then
            If DAType = "advance" Then
                ActiveSheet.Cells(12, 3 + deferredStart).Value = ActiveSheet.Range("B5").Value
                For i = 1 To (lengthSize - deferredStart - 1)
                growthFactor = (1 + ActiveSheet.Range("E5").Value) ^ i
                ActiveSheet.Cells(12, 3 + deferredStart + i).Value = ActiveSheet.Range("B5").Value * growthFactor
                Next i
            ElseIf DAType = "arrear" Then
                ActiveSheet.Cells(12, 4 + deferredStart).Value = ActiveSheet.Range("B5").Value
                For i = 1 To (lengthSize - deferredStart - 1)
                growthFactor = (1 + ActiveSheet.Range("E5").Value) ^ i
                ActiveSheet.Cells(12, 4 + deferredStart + i).Value = ActiveSheet.Range("B5").Value * growthFactor
                Next i
            End If
        Else
            If DAType = "advance" Then
                ActiveSheet.Cells(12, 3 + deferredStart).Value = ActiveSheet.Range("B5").Value
                For i = 1 To (lengthSize - deferredStart - 1)
                degrowthFactor = (1 - ActiveSheet.Range("F5").Value) ^ i
                ActiveSheet.Cells(12, 3 + deferredStart + i).Value = ActiveSheet.Range("B5").Value * degrowthFactor
                Next i
            ElseIf DAType = "arrear" Then
                ActiveSheet.Cells(12, 4 + deferredStart).Value = ActiveSheet.Range("B5").Value
                For i = 1 To (lengthSize - deferredStart - 1)
                degrowthFactor = (1 - ActiveSheet.Range("F5").Value) ^ i
                ActiveSheet.Cells(12, 4 + deferredStart + i).Value = ActiveSheet.Range("B5").Value * degrowthFactor
                Next i
            End If
        End If
    End If
End If
End Sub

Sub valueForm()
Dim pvOrfv As String
Dim dOrffactor As Double
Dim fPower As Long
Dim lengthSize As Long
Dim deferredStart As Integer
deferredStart = ActiveSheet.Range("D3").Value
pvOrfv = ActiveSheet.Range("B2").Value
If Sheets("Calculate").Range("B6").Value = "annually" Then
        lengthSize = 1 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "semiannually" Then
        lengthSize = 2 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "quarterly" Then
        lengthSize = 4 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "monthly" Then
       lengthSize = 12 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "daily" Then
        lengthSize = 365 * ActiveSheet.Range("B9").Value
    End If
If ActiveSheet.Range("B3").Value = "advance" And ActiveSheet.Range("C3").Value = "n" Then
    If pvOrfv = "present value" Then
        ActiveSheet.Range("B16").FormulaR1C1 = ""
        ActiveSheet.Range("C13").Value = ActiveSheet.Range("B5").Value
        For i = 1 To (lengthSize - 1)
        dOrffactor = (1 + ActiveSheet.Range("B8").Value) ^ -ActiveSheet.Cells(11, 3 + i).Value
        ActiveSheet.Cells(13, 3 + i).Value = ActiveSheet.Cells(12, 3 + i).Value * dOrffactor
        Next i
    ElseIf pvOrfv = "future value" Then
         For i = 1 To (lengthSize - 1)
         ActiveSheet.Range("C13").Value = ActiveSheet.Range("B5").Value * (1 + ActiveSheet.Range("B8").Value) ^ ActiveSheet.Cells(11, 3 + lengthSize).Value
        fPower = ActiveSheet.Cells(11, 3 + lengthSize).Value - ActiveSheet.Cells(11, 3 + i).Value
         dOrffactor = (1 + ActiveSheet.Range("B8").Value) ^ fPower
        ActiveSheet.Cells(13, 3 + i).Value = ActiveSheet.Cells(12, 3 + i).Value * dOrffactor
        Next i
    End If
ElseIf ActiveSheet.Range("B3").Value = "advance" And ActiveSheet.Range("C3").Value = "y" Then
    If pvOrfv = "present value" Then
        ActiveSheet.Range("B16").FormulaR1C1 = ""
        ActiveSheet.Cells(13, 3 + deferredStart).Value = ActiveSheet.Range("B5").Value * (1 + ActiveSheet.Range("B8").Value) ^ -deferredStart
        For i = 1 To (lengthSize - deferredStart - 1)
        dOrffactor = (1 + ActiveSheet.Range("B8").Value) ^ -ActiveSheet.Cells(11, 3 + deferredStart + i).Value
        ActiveSheet.Cells(13, 3 + deferredStart + i).Value = ActiveSheet.Cells(12, 3 + deferredStart + i).Value * dOrffactor
        Next i
    ElseIf pvOrfv = "future value" Then
         For i = 1 To (lengthSize - deferredStart - 1)
         ActiveSheet.Cells(13, 3 + deferredStart).Value = ActiveSheet.Range("B5").Value * (1 + ActiveSheet.Range("B8").Value) ^ (lengthSize - deferredStart)
        fPower = ActiveSheet.Cells(11, 3 + lengthSize).Value - ActiveSheet.Cells(11, 3 + deferredStart + i).Value
         dOrffactor = (1 + ActiveSheet.Range("B8").Value) ^ fPower
        ActiveSheet.Cells(13, 3 + deferredStart + i).Value = ActiveSheet.Cells(12, 3 + deferredStart + i).Value * dOrffactor
        Next i
    End If
ElseIf ActiveSheet.Range("B3").Value = "arrear" And ActiveSheet.Range("C3").Value = "n" Then
    If pvOrfv = "present value" Then
        ActiveSheet.Range("B16").FormulaR1C1 = ""
        For i = 1 To lengthSize
        dOrffactor = (1 + ActiveSheet.Range("B8").Value) ^ -ActiveSheet.Cells(11, 3 + i).Value
        ActiveSheet.Cells(13, 3 + i).Value = ActiveSheet.Cells(12, 3 + i).Value * dOrffactor
        Next i
    ElseIf pvOrfv = "future value" Then
         For i = 1 To lengthSize
        fPower = ActiveSheet.Cells(11, 3 + lengthSize).Value - ActiveSheet.Cells(11, 3 + i).Value
         dOrffactor = (1 + ActiveSheet.Range("B8").Value) ^ fPower
        ActiveSheet.Cells(13, 3 + i).Value = ActiveSheet.Cells(12, 3 + i).Value * dOrffactor
        Next i
    End If
ElseIf ActiveSheet.Range("B3").Value = "arrear" And ActiveSheet.Range("C3").Value = "y" Then
    If pvOrfv = "present value" Then
        ActiveSheet.Range("B16").FormulaR1C1 = ""
        For i = 1 To (lengthSize - deferredStart)
        dOrffactor = (1 + ActiveSheet.Range("B8").Value) ^ -ActiveSheet.Cells(11, 3 + deferredStart + i).Value
        ActiveSheet.Cells(13, 3 + deferredStart + i).Value = ActiveSheet.Cells(12, 3 + deferredStart + i).Value * dOrffactor
        Next i
    ElseIf pvOrfv = "future value" Then
         For i = 1 To (lengthSize - deferredStart)
        fPower = ActiveSheet.Cells(11, 3 + lengthSize).Value - ActiveSheet.Cells(11, 3 + deferredStart + i).Value
         dOrffactor = (1 + ActiveSheet.Range("B8").Value) ^ fPower
        ActiveSheet.Cells(13, 3 + deferredStart + i).Value = ActiveSheet.Cells(12, 3 + deferredStart + i).Value * dOrffactor
        Next i
    End If
End If
End Sub
Sub totalValue()
Dim sumall As Long
Dim lengthSize As Long
Dim rangeSize As Range
If Sheets("Calculate").Range("B6").Value = "annually" Then
        lengthSize = 1 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "semiannually" Then
        lengthSize = 2 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "quarterly" Then
        lengthSize = 4 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "monthly" Then
       lengthSize = 12 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "daily" Then
        lengthSize = 365 * ActiveSheet.Range("B9").Value
    End If
Set rangeSize = ActiveSheet.Range(Range("C13"), Range("C13").Offset(, lengthSize))
sumall = WorksheetFunction.Sum(rangeSize)
ActiveSheet.Range("B14").Value = sumall
ActiveSheet.Range("B14").Font.Bold = True
ActiveSheet.Range("B14").Font.Size = 15
End Sub
Sub netReturn()
Dim sumall As Double
Dim rangeSize As Range
Dim returnNet As Double
Dim lengthSize As Long
If ActiveSheet.Range("B2") = "future value" Then
    If Sheets("Calculate").Range("B6").Value = "annually" Then
        lengthSize = 1 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "semiannually" Then
        lengthSize = 2 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "quarterly" Then
        lengthSize = 4 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "monthly" Then
       lengthSize = 12 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "daily" Then
        lengthSize = 365 * ActiveSheet.Range("B9").Value
    End If
    Set rangeSize = ActiveSheet.Range(Range("C12"), Range("C12").Offset(, lengthSize))
    sumall = WorksheetFunction.Sum(rangeSize)
    returnNet = ActiveSheet.Range("B14").Value - sumall
    ActiveSheet.Range("B16").Value = returnNet
    ActiveSheet.Range("B16").Font.Bold = True
    ActiveSheet.Range("B16").Font.Size = 15
End If
End Sub
Sub totalInvest()
Dim lengthSize As Long
Dim rangeSize As Range
If Sheets("Calculate").Range("B6").Value = "annually" Then
        lengthSize = 1 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "semiannually" Then
        lengthSize = 2 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "quarterly" Then
        lengthSize = 4 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "monthly" Then
       lengthSize = 12 * ActiveSheet.Range("B9").Value
    ElseIf Sheets("Calculate").Range("B6").Value = "daily" Then
        lengthSize = 365 * ActiveSheet.Range("B9").Value
    End If
Set rangeSize = ActiveSheet.Range(Range("C12"), Range("C12").Offset(, lengthSize))
ActiveSheet.Range("B15").Value = WorksheetFunction.Sum(rangeSize)
End Sub
Sub constantAnnuityCalculator()
Attribute constantAnnuityCalculator.VB_ProcData.VB_Invoke_Func = " \n14"
Call clearContents
Call lengthTime
Call payment
Call valueForm
Call totalValue
Call totalInvest
Call netReturn
Call summaryReport
End Sub
Sub linknomConverter()
Attribute linknomConverter.VB_ProcData.VB_Invoke_Func = " \n14"
Call nomRateConversion
Call linknumValue
End Sub
Sub linkeffConverter()
Attribute linkeffConverter.VB_ProcData.VB_Invoke_Func = " \n14"
Call effRateConversion
Call linkeffValue
End Sub
Sub nomRateConversion()
Dim periodicInterest As Double
Dim divide As Double
If Sheets("Calculate").Range("B6").Value = "annually" Then
        divide = 1
    ElseIf Sheets("Calculate").Range("B6").Value = "semiannually" Then
        divide = 2
    ElseIf Sheets("Calculate").Range("B6").Value = "quarterly" Then
        divide = 4
    ElseIf Sheets("Calculate").Range("B6").Value = "monthly" Then
        divide = 12
    ElseIf Sheets("Calculate").Range("B6").Value = "daily" Then
        divide = 365
    End If
If ActiveSheet.Range("H3").Value = "Effective Interest Rate - annually" Then
    ActiveSheet.Range("I3").Value = ActiveSheet.Range("I2").Value
Else
    ActiveSheet.Range("I3").Value = ActiveSheet.Range("I2").Value / divide
End If
End Sub
Sub linknumValue()
If ActiveSheet.Range("I3").Value <> 0 Then
    ActiveSheet.Range("B8").Value = ActiveSheet.Range("I3").Value
End If
End Sub
Sub effRateConversion()
 Dim periodicFactor As Double
    Dim divide As Double
    
    If Sheets("Calculate").Range("B6").Value = "annually" Then
        divide = 1 / 1
    ElseIf Sheets("Calculate").Range("B6").Value = "semiannually" Then
        divide = 1 / 2
    ElseIf Sheets("Calculate").Range("B6").Value = "quarterly" Then
        divide = 1 / 4
    ElseIf Sheets("Calculate").Range("B6").Value = "monthly" Then
        divide = 1 / 12
    ElseIf Sheets("Calculate").Range("B6").Value = "daily" Then
        divide = 1 / 365
    End If
periodicFactor = (1 + ActiveSheet.Range("I5").Value) ^ divide
    ActiveSheet.Range("I6").Value = periodicFactor - 1
End Sub
Sub linkeffValue()
If ActiveSheet.Range("I6").Value <> 0 Then
    ActiveSheet.Range("B8").Value = ActiveSheet.Range("I6").Value
End If
End Sub
Sub clearContents()
If ActiveSheet.Range("C12") <> 0 Then
    Range("C12:C14").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.clearContents
Else
    Range("D12:D14").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.clearContents
        Sheets("Calculate").Range("E15:E17").clearContents
End If
Sheets("Calculate").Range("E1").Select
End Sub
Sub sheetName()
Dim nextSheet As Worksheet
ActiveSheet.Name = "Summary_Report"
Set nextSheet = Sheets.Add(After:=ActiveSheet)
nextSheet.Name = "Calculate"
End Sub

Sub calTemplate()
Attribute calTemplate.VB_ProcData.VB_Invoke_Func = "T\n14"
Call sheetName
ActiveSheet.Range("A1").Value = "Discrete Annuity Calculator"
With ActiveSheet.Range("A1")
    .Font.Bold = True
    .Font.Size = 18
    .Font.Name = "ADLaM Display"
    .Font.Strikethrough = False
    .Font.Superscript = False
    .Font.Subscript = False
    .Font.OutlineFont = False
    .Font.Shadow = False
    .Font.Underline = xlUnderlineStyleNone
    .Font.ThemeColor = xlThemeColorLight1
    .Font.TintAndShade = 0
    .Font.ThemeFont = xlThemeFontNone
End With
ActiveSheet.Columns("A").EntireColumn.AutoFit
ActiveSheet.Range("A2").Value = "Value Form"
ActiveSheet.Range("A3").Value = "Advance or Arrear"
ActiveSheet.Range("A5").Value = "Payment & Varryng Amount"
ActiveSheet.Range("B5").NumberFormat = "#,##0"
ActiveSheet.Range("A6").Value = "Time Frame"
ActiveSheet.Range("A8").Value = "=""Effective Interest Rate - "" " & "&R[-2]C[1]"
ActiveSheet.Range("B8").NumberFormat = "0.00%"
ActiveSheet.Range("A9").Value = "Time Horizon"
    With ActiveSheet.Range("A4:B4").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    With ActiveSheet.Range("A2:A9").Font
        .Bold = True
        .Name = "Agency FB"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    ActiveSheet.Range("B1").Value = "Input"
    With ActiveSheet.Range("B1").Font
        .Bold = True
        .Name = "Agency FB"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With ActiveSheet.Range("B1").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    ActiveSheet.Columns("B").ColumnWidth = 20
    ActiveSheet.Range("C2").Value = "Deferred(y/n)"
    ActiveSheet.Range("D2").Value = "Period"
    With ActiveSheet.Range("C2:D2").Font
        .Bold = True
        .Name = "Agency FB"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
With Sheets("Calculate").Range("A7:B7").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
ActiveSheet.Range("C4").Value = "Increasing Amount"
ActiveSheet.Columns("C").EntireColumn.AutoFit
ActiveSheet.Range("C5").NumberFormat = "#,##0.00"
ActiveSheet.Range("D4").Value = "Decresing Amount"
ActiveSheet.Columns("D").EntireColumn.AutoFit
ActiveSheet.Range("D5").NumberFormat = "#,##0.00"
ActiveSheet.Range("E4").Value = "Positive Growth"
ActiveSheet.Columns("E").EntireColumn.AutoFit
ActiveSheet.Range("E5").NumberFormat = "0.00%"
ActiveSheet.Range("F4").Value = "Negative Growth"
ActiveSheet.Columns("F").EntireColumn.AutoFit
ActiveSheet.Range("F5").NumberFormat = "0.00%"
    With ActiveSheet.Range("C4:F4").Font
        .Bold = True
        .Name = "Agency FB"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With ActiveSheet.Range("C4").Font
        .Color = -10477568
        .TintAndShade = 0
    End With
    With ActiveSheet.Range("D4").Font
        .Color = -16777024
        .TintAndShade = 0
    End With
    With ActiveSheet.Range("E4").Font
        .Color = -10477568
        .TintAndShade = 0
    End With
    With ActiveSheet.Range("F4").Font
        .Color = -16777024
        .TintAndShade = 0
    End With
ActiveSheet.Range("B1:B9").Select
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
   Selection.Borders(xlInsideVertical).LineStyle = xlNone
   Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
ActiveSheet.Range("H1").Value = "Rate Conversion Calculator"
    With ActiveSheet.Range("H1")
        .Font.Bold = True
        .Font.Size = 18
        .Font.Name = "ADLaM Display"
        .Font.Strikethrough = False
        .Font.Superscript = False
        .Font.Subscript = False
        .Font.OutlineFont = False
        .Font.Shadow = False
        .Font.Underline = xlUnderlineStyleNone
        .Font.ThemeColor = xlThemeColorLight1
        .Font.TintAndShade = 0
        .Font.ThemeFont = xlThemeFontNone
    End With
ActiveSheet.Range("H1:L1").Merge
ActiveSheet.Columns("H").ColumnWidth = 20.18
ActiveSheet.Range("H2").Value = "Nominal Interest"
ActiveSheet.Range("I2").NumberFormat = "0.00%"
ActiveSheet.Range("H3").Value = "=""Effective Interest Rate - ""&R[3]C[-6]"
ActiveSheet.Range("I3").NumberFormat = "0.00%"
ActiveSheet.Range("H5").Value = "Annual Effective Interest"
ActiveSheet.Range("I5").NumberFormat = "0.00%"
ActiveSheet.Range("H6").Value = "=""Effective Interest Rate - ""&R[0]C[-6]"
ActiveSheet.Range("I6").NumberFormat = "0.00%"
    With ActiveSheet.Range("H2:H6").Font
        .Bold = True
        .Name = "Agency FB"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
ActiveSheet.Range("I2:I3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    With ActiveSheet.Range("I5:I6").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
ActiveSheet.Range("H1:I7").Select
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
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
ActiveSheet.Range("A11").Value = "Length Of Time"
ActiveSheet.Range("A12").Value = "Payement"
ActiveSheet.Range("A13").Value = "Value Form"
ActiveSheet.Range("A14").Value = "=""Total - "" " & "&R[-12]C[1]"
ActiveSheet.Range("B14").NumberFormat = "#,##0.00"
ActiveSheet.Range("A15").Value = "Total Money Invested"
ActiveSheet.Range("B15").NumberFormat = "#,##0.00"
ActiveSheet.Range("B15").Select
    With Selection.Font
        .Bold = True
        .Size = 14
    End With
ActiveSheet.Range("A15").Select
        With Selection.Font
            .Bold = True
            .Name = "Abadi"
            .Size = 13
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontNone
        End With
    With ActiveSheet.Range("A11:A14").Font
        .Bold = True
        .Name = "Agency FB"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
ActiveSheet.Range("A14").Font.Size = 15
Range("A15:B15").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13434828
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
Range("B12:B13").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.NumberFormat = "#,##0.00"
ActiveSheet.Range("A16").Value = "Net Return On future value"
ActiveSheet.Range("B16").NumberFormat = "#,##0.00"
ActiveSheet.Range("A16").Select
        With Selection.Font
            .Name = "Abadi"
            .Size = 16
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontNone
        End With
        With Selection.Font
            .Name = "Abadi"
            .Size = 16
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontNone
        End With
        Selection.Font.Bold = True
ActiveSheet.Range("A16:B16").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
ActiveSheet.Range("A11:A14").Select
    Range(Selection, Selection.End(xlToRight)).Select
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
Range("B2").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="present value, future value"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
Range("C3").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="y, n"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
Range("B3").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="advance, arrear"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
Range("B6").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="annually, semiannually, quarterly, monthly, daily"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
Sheets("Calculate").Range("D17").FormulaR1C1 = "=""-"" &  R[-15]C[-2]"
Sheets("Calculate").Range("E16:E17").Select
    Selection.NumberFormat = "0.00"
    Selection.Style = "Comma"
    Selection.Font.Bold = True
Sheets("Calculate").Range("D15:D17").Select
    With Selection.Font
        .Name = "Agency FB"
        .Size = 12
        .Bold = True
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
Sheets("Calculate").Range("D15").Value = "Search Cashflow[]"
Sheets("Calculate").Range("D16").Value = "Result"
Sheets("Calculate").Range("E15:E17").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
 Sheets("Calculate").Shapes.AddShape(msoShapeRectangle, 376, 3, 85, 21).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Run Me-PV/FV"
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 6). _
        ParagraphFormat
        .FirstLineIndent = 0
        .Alignment = msoAlignLeft
    End With
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 6).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 11
        .Name = "+mn-lt"
    End With
  Sheets("Calculate").Shapes.Range(Array("Rectangle 1")).Select
    Selection.OnAction = "PERSONAL.XLSB!constantAnnuityCalculator"
    With Selection.ShapeRange.ThreeD
        .SetPresetCamera (msoCameraOrthographicFront)
        .RotationX = 0
        .RotationY = 0
        .RotationZ = 0
        .FieldOfView = 0
        .LightAngle = 0
        .PresetLighting = msoLightRigSoft
        .PresetMaterial = msoMaterialMatte2
        .Depth = 0
        .ContourWidth = 3.5
        .ContourColor.RGB = RGB(255, 255, 255)
        .BevelTopType = msoBevelArtDeco
        .BevelTopInset = 5
        .BevelTopDepth = 5
        .BevelBottomType = msoBevelNone
    End With
    With Selection.ShapeRange.Shadow
        .Type = msoShadow25
        .Visible = msoTrue
        .Style = msoShadowStyleOuterShadow
        .Blur = 8.5
        .OffsetX = 6.1232339957E-17
        .OffsetY = 1
        .RotateWithShape = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
        .Size = 100
    End With
    Selection.ShapeRange.Line.Visible = msoFalse
'Button for the nominal interest converter calculator
 Sheets("Calculate").Shapes.AddShape(msoShapeRectangle, 918, 30.5, 36.5, 15.5).Select
    With Selection.ShapeRange.ThreeD
        .SetPresetCamera (msoCameraOrthographicFront)
        .RotationX = 0
        .RotationY = 0
        .RotationZ = 0
        .FieldOfView = 0
        .LightAngle = 0
        .PresetLighting = msoLightRigSoft
        .PresetMaterial = msoMaterialMatte2
        .Depth = 0
        .ContourWidth = 3.5
        .ContourColor.RGB = RGB(255, 255, 255)
        .BevelTopType = msoBevelArtDeco
        .BevelTopInset = 5
        .BevelTopDepth = 5
        .BevelBottomType = msoBevelNone
    End With
    With Selection.ShapeRange.Shadow
        .Type = msoShadow25
        .Visible = msoTrue
        .Style = msoShadowStyleOuterShadow
        .Blur = 8.5
        .OffsetX = 6.1232339957E-17
        .OffsetY = 1
        .RotateWithShape = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
        .Size = 100
    End With
    Selection.ShapeRange.Line.Visible = msoFalse
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Run Me"
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 6). _
        ParagraphFormat
        .FirstLineIndent = 0
        .Alignment = msoAlignLeft
    End With
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 6).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 7
        .Name = "+mn-lt"
    End With
     Sheets("Calculate").Shapes.Range(Array("Rectangle 2")).Select
    Selection.OnAction = "PERSONAL.XLSB!linknomConverter"
'Button for the effective interet rate converter
Sheets("Calculate").Shapes.AddShape(msoShapeRectangle, 919, 76.5, 36, 13.5).Select
    With Selection.ShapeRange.ThreeD
        .SetPresetCamera (msoCameraOrthographicFront)
        .RotationX = 0
        .RotationY = 0
        .RotationZ = 0
        .FieldOfView = 0
        .LightAngle = 0
        .PresetLighting = msoLightRigSoft
        .PresetMaterial = msoMaterialMatte2
        .Depth = 0
        .ContourWidth = 3.5
        .ContourColor.RGB = RGB(255, 255, 255)
        .BevelTopType = msoBevelArtDeco
        .BevelTopInset = 5
        .BevelTopDepth = 5
        .BevelBottomType = msoBevelNone
    End With
    With Selection.ShapeRange.Shadow
        .Type = msoShadow25
        .Visible = msoTrue
        .Style = msoShadowStyleOuterShadow
        .Blur = 8.5
        .OffsetX = 6.1232339957E-17
        .OffsetY = 1
        .RotateWithShape = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
        .Size = 100
    End With
    Selection.ShapeRange.Line.Visible = msoFalse
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Run Me"
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 6). _
        ParagraphFormat
        .FirstLineIndent = 0
        .Alignment = msoAlignLeft
    End With
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 6).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 7
        .Name = "+mn-lt"
    End With
    Sheets("Calculate").Shapes.Range(Array("Rectangle 3")).Select
    Selection.OnAction = "PERSONAL.XLSB!linkeffConverter"

 Sheets("Calculate").Shapes.AddShape(msoShapeRectangle, 628, 238, 49.5, 16).Select
    With Selection.ShapeRange.ThreeD
        .SetPresetCamera (msoCameraOrthographicFront)
        .RotationX = 0
        .RotationY = 0
        .RotationZ = 0
        .FieldOfView = 0
        .LightAngle = 0
        .PresetLighting = msoLightRigSoft
        .PresetMaterial = msoMaterialMatte2
        .Depth = 0
        .ContourWidth = 3.5
        .ContourColor.RGB = RGB(255, 255, 255)
        .BevelTopType = msoBevelArtDeco
        .BevelTopInset = 5
        .BevelTopDepth = 5
        .BevelBottomType = msoBevelNone
    End With
    With Selection.ShapeRange.Shadow
        .Type = msoShadow25
        .Visible = msoTrue
        .Style = msoShadowStyleOuterShadow
        .Blur = 8.5
        .OffsetX = 6.1232339957E-17
        .OffsetY = 1
        .RotateWithShape = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
        .Size = 100
    End With
    Selection.ShapeRange.Line.Visible = msoFalse
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Run Me"
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 6). _
        ParagraphFormat
        .FirstLineIndent = 0
        .Alignment = msoAlignLeft
    End With
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 6).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 11
        .Name = "+mn-lt"
    End With
    Sheets("Calculate").Shapes.Range(Array("Rectangle 4")).Select
    Selection.OnAction = "PERSONAL.XLSB!searchCashflow"
Sheets("Calculate").Range("C3").Value = "n"
Sheets("Calculate").Range("E1").Select
MsgBox "This calculator simplifies managing cashflows, making financial tasks quicker and easier." & vbCrLf & "Developer: Sekou M. Kamara", vbOKOnly + vbInformation, "Cashflow"
End Sub
Sub summaryReport()
Sheets("Summary_Report").Range("B3:B5").clearContents
Sheets("Summary_Report").Columns("A").ColumnWidth = 30.45
Sheets("Summary_Report").Range("A1").Value = "Summary Report"
With Sheets("Summary_Report").Range("A1").Font
    .Name = "ADLaM Display"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
With Sheets("Summary_Report").Range("A1").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.749992370372631
        .PatternTintAndShade = 0
    End With
If Sheets("Calculate").Range("B2").Value = "present value" Then
    Sheets("Summary_Report").Range("A3").Value = "Total Present Value"
ElseIf Sheets("Calculate").Range("B2").Value = "future value" Then
    Sheets("Summary_Report").Range("A3").Value = "Total Future Value"
End If
Sheets("Summary_Report").Range("A4").Value = "Total Money Invested"
Sheets("Summary_Report").Range("A5").Value = "Total Return on Investment"

With Sheets("Summary_Report").Range("A3:A5").Font
        .Name = "ADLaM Display"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Sheets("Summary_Report").Range("B3:B5").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
Sheets("Summary_Report").Range("B3").FormulaR1C1 = "=Calculate!R[11]C"
Sheets("Summary_Report").Range("B4").FormulaR1C1 = "=Calculate!R[11]C"
If Sheets("Calculate").Range("B2").Value = "future value" Then
 Sheets("Summary_Report").Range("B5").FormulaR1C1 = "=Calculate!R[11]C"
ElseIf Sheets("Calculate").Range("B2").Value = "present value" Then
    Sheets("Summary_Report").Range("B5").Value = "N/A"
End If
With Sheets("Summary_Report").Range("B3:B5").Font
        .Name = "Aptos Narrow"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
Sheets("Calculate").Range("E1").Select
End Sub
Sub yearEquivalence()
If Sheet("Calculate").Range("B6").Value = "annually" Then
    Sheet("Calculate").Range("B7").Value = 1
ElseIf Sheet("Calculate").Range("B6").Value = "semiannually" Then
    Sheet("Calculate").Range("B7").Value = 2
ElseIf Sheet("Calculate").Range("B6").Value = "quarterly" Then
    Sheet("Calculate").Range("B7").Value = 4
ElseIf Sheet("Calculate").Range("B7").Value = "monthly" Then
    Sheet("Calculate").Range("B7").Value = 12
ElseIf Sheet("Calculate").Range("B6").Value = "daily" Then
    Sheet("Calculate").Range("B7").Value = 365
End Sub
Sub searchCashflow()
Attribute searchCashflow.VB_ProcData.VB_Invoke_Func = " \n14"
Dim indexSearch As Integer
Dim indexPV As Double
indexSearch = Sheets("Calculate").Range("E15").Value
If indexSearch > 0 Then
Sheets("Calculate").Range("E16").Value = Sheets("Calculate").Cells(12, 3 + indexSearch).Value
Sheets("Calculate").Range("E17").Value = Sheets("Calculate").Cells(13, 3 + indexSearch).Value
End If
End Sub



