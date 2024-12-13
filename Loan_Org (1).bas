Attribute VB_Name = "Module2"
Dim i As Integer
Dim frameEquivalent As Integer
Dim timeLength As Integer


Sub timeHorizon()
If Sheets("Loan_Schedule").Range("B4").Value = "annually" Then
    frameEquivalent = 1
ElseIf Sheets("Loan_Schedule").Range("B4").Value = "semi annually" Then
    frameEquivalent = 2
ElseIf Sheets("Loan_Schedule").Range("B4").Value = "quarterly" Then
    frameEquivalent = 4
ElseIf Sheets("Loan_Schedule").Range("B4").Value = "monthly" Then
    frameEquivalent = 12
ElseIf Sheets("Loan_Schedule").Range("B4").Value = "daily" Then
    frameEquivalent = 365
End If
timeLength = frameEquivalent * Sheets("Loan_Schedule").Range("B5").Value
For i = 1 To timeLength
    Sheets("Loan_Schedule").Cells(11 + i, 1).Value = i
Next
End Sub
Sub repaymentAmount()
Dim divisor As Double
Dim repay As Double
Dim interestPay As Double
Dim growthFactor As Double
Dim levelFactor As Double
Dim upperFactor As Double
Dim upperFactorB As Double
Dim upperFactorC As Double
Dim interest  As Double
Dim initialAmount As Double
Dim amountBorrow As Double
Dim increAmount As Double
Dim lumpSum As Double

Call timeHorizon

Sheets("Loan_Schedule").Cells(12, 2).Value = Sheets("Loan_Schedule").Range("B3").Value
If Sheets("Loan_Schedule").Range("B2").Value = "amortized" Then
    If Sheets("Loan_Schedule").Range("B7").Value = "level" Then
        divisor = 1 - (1 + Sheets("Loan_Schedule").Range("B8").Value) ^ -timeLength
        repay = (Sheets("Loan_Schedule").Range("B8").Value * Sheets("Loan_Schedule").Range("B3").Value) / divisor
        For i = 1 To timeLength
        Sheets("Loan_Schedule").Cells(11 + i, 3).Value = repay
        Next
    ElseIf Sheets("Loan_Schedule").Range("B7").Value = "increasing" Then
        increAmount = Sheets("Loan_Schedule").Range("C7").Value
        amountBorrow = Sheets("Loan_Schedule").Range("B3").Value
       'took off numperiod and replace with timeLength in hare
        interest = Sheets("Loan_Schedule").Range("B8").Value
        upperFactor = 1 - (1 + interest) ^ -timeLength
        levelFactor = upperFactor / interest
        upperFactorB = levelFactor * (1 + interest) - (timeLength * (1 + interest) ^ -timeLength)
        growthFactor = upperFactorB / interest
        upperFactorC = amountBorrow - (increAmount * (growthFactor - levelFactor))
        initialAmount = upperFactorC / levelFactor
        For i = 1 To timeLength
        Sheets("Loan_Schedule").Cells(11 + i, 3).Value = initialAmount + (i - 1) * increAmount
        Next
    ElseIf Sheets("Loan_Schedule").Range("B7").Value = "decreasing" Then
        increAmount = -Sheets("Loan_Schedule").Range("C7").Value
        amountBorrow = Sheets("Loan_Schedule").Range("B3").Value
       ' numperiod was remove from here inplace of timeLength
        interest = Sheets("Loan_Schedule").Range("B8").Value
        upperFactor = 1 - (1 + interest) ^ -timeLength
        levelFactor = upperFactor / interest
        upperFactorB = levelFactor * (1 + interest) - (timeLength * (1 + interest) ^ -timeLength)
        growthFactor = upperFactorB / interest
        upperFactorC = amountBorrow - (increAmount * (growthFactor - levelFactor))
        initialAmount = upperFactorC / levelFactor
        For i = 1 To timeLength
        Sheets("Loan_Schedule").Cells(11 + i, 3).Value = initialAmount + (i - 1) * increAmount
        Next
    End If
    'These are the four  major lines execuiting the other major areas with the exception of repayment at index 3
    For i = 1 To timeLength
        Sheets("Loan_Schedule").Cells(11 + i, 4).Value = Sheets("Loan_Schedule").Range("B8").Value * Sheets("Loan_Schedule").Cells(11 + i, 2).Value
        Sheets("Loan_Schedule").Cells(11 + i, 5).Value = Sheets("Loan_Schedule").Cells(11 + i, 3).Value - Sheets("Loan_Schedule").Cells(11 + i, 4).Value
        Sheets("Loan_Schedule").Cells(12 + i, 2).Value = Sheets("Loan_Schedule").Cells(11 + i, 2).Value - Sheets("Loan_Schedule").Cells(11 + i, 5).Value
        Sheets("Loan_Schedule").Cells(11 + i, 6).Value = Sheets("Loan_Schedule").Cells(12 + i, 2).Value
    Next
    'introduce timeLength in clear content for excess running of code
    Sheets("Loan_Schedule").Cells(12 + timeLength, 2).clearContents
    Sheets("Loan_Schedule").Range("D1").Select
ElseIf Sheets("Loan_Schedule").Range("B2").Value = "interest only" Then
     For i = 1 To timeLength
        Sheets("Loan_Schedule").Cells(11 + i, 2) = Sheets("Loan_Schedule").Range("B3").Value
        Sheets("Loan_Schedule").Cells(11 + i, 4) = Sheets("Loan_Schedule").Cells(11 + i, 2).Value * Sheets("Loan_Schedule").Range("B8").Value
        If i <> timeLength Then
            Sheets("Loan_Schedule").Cells(11 + i, 6) = Sheets("Loan_Schedule").Cells(11 + i, 2)
            Sheets("Loan_Schedule").Cells(11 + i, 3) = Sheets("Loan_Schedule").Cells(11 + i, 4).Value
            Sheets("Loan_Schedule").Cells(11 + i, 5) = 0
        Else
            Sheets("Loan_Schedule").Cells(11 + i, 6) = 0
            Sheets("Loan_Schedule").Cells(11 + i, 3) = Sheets("Loan_Schedule").Range("B3").Value + Sheets("Loan_Schedule").Cells(11 + i, 4).Value
            Sheets("Loan_Schedule").Cells(11 + i, 5) = Sheets("Loan_Schedule").Range("B3").Value
        End If
    Next
ElseIf Sheets("Loan_Schedule").Range("B2").Value = "partially amortized" Then
    If Sheets("Loan_Schedule").Range("B7").Value = "level" Then
        lumpSum = Sheets("Loan_Schedule").Range("C2").Value
       'remove numperiod from here inplace of timeLength
        interest = Sheets("Loan_Schedule").Range("B8").Value
        upperFactor = 1 - (1 + interest) ^ -timeLength
        levelFactor = upperFactor / interest
        repay = (Sheets("Loan_Schedule").Range("B3").Value - Sheets("Loan_Schedule").Range("C2").Value * (1 + interest) ^ -timeLength) / levelFactor
        For i = 1 To timeLength
        If i <> timeLength Then
        Sheets("Loan_Schedule").Cells(11 + i, 3).Value = repay
        Else
        Sheets("Loan_Schedule").Cells(11 + i, 3).Value = repay + lumpSum
        End If
        Next
         'These are the four  major lines execuiting the other major areas with the exception of repayment at index 3
        For i = 1 To timeLength
            Sheets("Loan_Schedule").Cells(11 + i, 4).Value = Sheets("Loan_Schedule").Range("B8").Value * Sheets("Loan_Schedule").Cells(11 + i, 2).Value
            Sheets("Loan_Schedule").Cells(11 + i, 5).Value = Sheets("Loan_Schedule").Cells(11 + i, 3).Value - Sheets("Loan_Schedule").Cells(11 + i, 4).Value
            Sheets("Loan_Schedule").Cells(12 + i, 2).Value = Sheets("Loan_Schedule").Cells(11 + i, 2).Value - Sheets("Loan_Schedule").Cells(11 + i, 5).Value
            Sheets("Loan_Schedule").Cells(11 + i, 6).Value = Sheets("Loan_Schedule").Cells(12 + i, 2).Value
        Next
        'put timeLength in clear content
        Sheets("Loan_Schedule").Cells(12 + timeLength, 2).clearContents
        Sheets("Loan_Schedule").Range("D1").Select
    ElseIf Sheets("Loan_Schedule").Range("B7").Value = "increasing" Then
        lumpSum = Sheets("Loan_Schedule").Range("C2").Value
        increAmount = Sheets("Loan_Schedule").Range("C7").Value
        amountBorrow = Sheets("Loan_Schedule").Range("B3").Value
        'remove numperiod from here inplace of timeLength
        interest = Sheets("Loan_Schedule").Range("B8").Value
        upperFactor = 1 - (1 + interest) ^ -timeLength
        levelFactor = upperFactor / interest
        upperFactorB = levelFactor * (1 + interest) - (timeLength * (1 + interest) ^ -timeLength)
        growthFactor = upperFactorB / interest
        initialAmount = (Sheets("Loan_Schedule").Range("B3").Value + increAmount * (levelFactor - growthFactor) - lumpSum * (1 + interest) ^ -timeLength) / levelFactor
        For i = 1 To timeLength
            If i <> timeLength Then
                Sheets("Loan_Schedule").Cells(11 + i, 3).Value = initialAmount + (i - 1) * increAmount
            Else
                Sheets("Loan_Schedule").Cells(11 + i, 3).Value = initialAmount + (i - 1) * increAmount + lumpSum
            End If
        Next
    ElseIf Sheets("Loan_Schedule").Range("B7").Value = "decreasing" Then
        lumpSum = Sheets("Loan_Schedule").Range("C2").Value
        increAmount = -Sheets("Loan_Schedule").Range("C7").Value
        amountBorrow = Sheets("Loan_Schedule").Range("B3").Value
        'remove numperiod inplace of  timeLength
        interest = Sheets("Loan_Schedule").Range("B8").Value
        upperFactor = 1 - (1 + interest) ^ -timeLength
        levelFactor = upperFactor / interest
        upperFactorB = levelFactor * (1 + interest) - (timeLength * (1 + interest) ^ -timeLength)
        growthFactor = upperFactorB / interest
        initialAmount = (Sheets("Loan_Schedule").Range("B3").Value + increAmount * (levelFactor - growthFactor) - lumpSum * (1 + interest) ^ -timeLength) / levelFactor
        For i = 1 To timeLength
            If i <> timeLength Then
                Sheets("Loan_Schedule").Cells(11 + i, 3).Value = initialAmount + (i - 1) * increAmount
            Else
                Sheets("Loan_Schedule").Cells(11 + i, 3).Value = initialAmount + (i - 1) * increAmount + lumpSum
            End If
        Next
    End If
     'These are the four  major lines execuiting the other major areas with the exception of repayment at index 3
        For i = 1 To timeLength
            Sheets("Loan_Schedule").Cells(11 + i, 4).Value = Sheets("Loan_Schedule").Range("B8").Value * Sheets("Loan_Schedule").Cells(11 + i, 2).Value
            Sheets("Loan_Schedule").Cells(11 + i, 5).Value = Sheets("Loan_Schedule").Cells(11 + i, 3).Value - Sheets("Loan_Schedule").Cells(11 + i, 4).Value
            Sheets("Loan_Schedule").Cells(12 + i, 2).Value = Sheets("Loan_Schedule").Cells(11 + i, 2).Value - Sheets("Loan_Schedule").Cells(11 + i, 5).Value
            Sheets("Loan_Schedule").Cells(11 + i, 6).Value = Sheets("Loan_Schedule").Cells(12 + i, 2).Value
        Next
        Sheets("Loan_Schedule").Cells(12 + timeLength, 2).clearContents
        Sheets("Loan_Schedule").Range("D1").Select
End If
Sheets("Loan_Schedule").Range("D10").FormulaR1C1 = "=SUM(R[2]C:R[" & (1 + timeLength) & "]C)"
End Sub

Sub effRateConversionLoan()
Dim periodicFactor As Double
Dim divide As Double
    
    If Sheets("Loan_Schedule").Range("B4").Value = "annually" Then
        divide = 1 / 1
    ElseIf Sheets("Loan_Schedule").Range("B4").Value = "semi annually" Then
        divide = 1 / 2
    ElseIf Sheets("Loan_Schedule").Range("B4").Value = "quarterly" Then
        divide = 1 / 4
    ElseIf Sheets("Loan_Schedule").Range("B4").Value = "monthly" Then
        divide = 1 / 12
    ElseIf Sheets("Loan_Schedule").Range("B4").Value = "daily" Then
        divide = 1 / 365
    End If
periodicFactor = (1 + Sheets("Loan_Schedule").Range("J5").Value) ^ divide
    Sheets("Loan_Schedule").Range("J6").Value = periodicFactor - 1

If Sheets("Loan_Schedule").Range("J6").Value <> 0 Then
    Sheets("Loan_Schedule").Range("B8").Value = Sheets("Loan_Schedule").Range("J6").Value
End If
End Sub
Sub nomRateConversionLoan()
Dim periodicInterest As Double
Dim divide As Double
If Sheets("Loan_Schedule").Range("B4").Value = "annually" Then
        divide = 1
    ElseIf Sheets("Loan_Schedule").Range("B4").Value = "semi annually" Then
        divide = 2
    ElseIf Sheets("Loan_Schedule").Range("B4").Value = "quarterly" Then
        divide = 4
    ElseIf Sheets("Loan_Schedule").Range("B4").Value = "monthly" Then
        divide = 12
    ElseIf Sheets("Loan_Schedule").Range("B4").Value = "daily" Then
        divide = 365
    End If
Sheets("Loan_Schedule").Range("J3").Value = Sheets("Loan_Schedule").Range("J2").Value / divide

If Sheets("Loan_Schedule").Range("J3").Value <> 0 Then
    Sheets("Loan_Schedule").Range("B8").Value = Sheets("Loan_Schedule").Range("J3").Value
End If
End Sub
Sub trackRepay()
Dim ballance As Double
If Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 8) < 1 Then
        ballance = Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 3).Value - Sheets("Loan_Schedule").Range("F2").Value
        If ballance < 1 Then
            Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 7).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = 0.499984740745262
                .PatternTintAndShade = 0
             End With
            Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 7).Font.Bold = True
            Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 7).Value = "Complete"
            Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 8).clearContents
            
        ElseIf ballance >= 1 And Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 7) <> "Complete" Then
            Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 7).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 7).Font.Bold = True
            Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 8).Select
             With Selection.Font
            .Color = -16776961
            .TintAndShade = 0
             End With
            Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 8).Font.Bold = True
            Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 7).Value = "incomplete"
            Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 8).Value = ballance
        End If
Else
    ballance = Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 8).Value - Sheets("Loan_Schedule").Range("F2").Value
    If ballance < 1 Then
            Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 7).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = 0.499984740745262
                .PatternTintAndShade = 0
             End With
            Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 7).Font.Bold = True
            Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 7).Value = "Complete"
            Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 8).clearContents
        Else
            Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 7).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 7).Font.Bold = True
            Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 8).Select
             With Selection.Font
            .Color = -16776961
            .TintAndShade = 0
             End With
            Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 8).Font.Bold = True
            Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 7).Value = "incomplete"
            Sheets("Loan_Schedule").Cells(11 + Sheets("Loan_Schedule").Range("G2").Value, 8).Value = ballance
        End If
End If
Sheets("Loan_Schedule").Range("D1").Select
End Sub
Sub tempLoan()
Attribute tempLoan.VB_ProcData.VB_Invoke_Func = "Y\n14"
ActiveSheet.Name = "Loan_Schedule"

Sheets("Loan_Schedule").Range("B2").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="interest only, amortized, partially amortized"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
Sheets("Loan_Schedule").Range("B7").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="level, increasing, decreasing"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
With Sheets("Loan_Schedule").Range("B4").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="annually, semi annually, quarterly, monthly, daily"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

Sheets("Loan_Schedule").Range("A1").ColumnWidth = 16
Sheets("Loan_Schedule").Range("C1").ColumnWidth = 14
Sheets("Loan_Schedule").Range("G11").ColumnWidth = 10
Sheets("Loan_Schedule").Range("H11").ColumnWidth = 10
Sheets("Loan_Schedule").Range("I1").ColumnWidth = 23
Sheets("Loan_Schedule").Range("A1:F1").ColumnWidth = 16

Sheets("Loan_Schedule").Range("F1:G2").Select
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

Sheets("Loan_Schedule").Range("B1:B8").Select
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

Sheets("Loan_Schedule").Range("I1:L1").Merge
Sheets("Loan_Schedule").Range("E1:E2").Merge
   With Sheets("Loan_Schedule").Range("I1").Font
        .Name = "ADLaM Display"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
        .Bold = True
    End With

Sheets("Loan_Schedule").Range("A11:F11").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
Sheets("Loan_Schedule").Range("A2:A8").Select
   With Selection.Font
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
        .Bold = True
    End With
    
Sheets("Loan_Schedule").Range("B1:C1").Select
    With Selection.Font
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
        .Bold = True
    End With
    
Sheets("Loan_Schedule").Range("C6").Select
    With Selection.Font
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
        .Bold = True
    End With
    
Sheets("Loan_Schedule").Range("F1:G1").Select
    With Selection.Font
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
        .Bold = True
    End With
    
Sheets("Loan_Schedule").Range("A11:F11").Select
    With Selection.Font
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
        .Bold = True
    End With
    
Sheets("Loan_Schedule").Range("F2").NumberFormat = "#,##0"
Sheets("Loan_Schedule").Range("B3").NumberFormat = "#,##0"
Sheets("Loan_Schedule").Range("C2").NumberFormat = "#,##0"
Sheets("Loan_Schedule").Range("C7").NumberFormat = "#,##0"
Sheets("Loan_Schedule").Range("B8").NumberFormat = "0.00%"
Sheets("Loan_Schedule").Range("J2").NumberFormat = "0.00%"
Sheets("Loan_Schedule").Range("J3").NumberFormat = "0.00%"
Sheets("Loan_Schedule").Range("J5").NumberFormat = "0.00%"
Sheets("Loan_Schedule").Range("J6").NumberFormat = "0.00%"
Sheets("Loan_Schedule").Range("H12:H1048576").NumberFormat = "#,##0"
Sheets("Loan_Schedule").Range("B12:F1048576").NumberFormat = "#,##0"

With Sheets("Loan_Schedule").Range("I2:I6").Font
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
Sheets("Loan_Schedule").Range("J2:J3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
With Sheets("Loan_Schedule").Range("J5:J6").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
Sheets("Loan_Schedule").Range("I1:J7").Select
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
    
With Sheets("Loan_Schedule").Range("B1").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
 With Sheets("Loan_Schedule").Range("D10").Font
        .Name = "Aptos Narrow"
        .Size = 14
        .Bold = True
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
    
Sheets("Loan_Schedule").Shapes.AddShape(msoShapeRectangle, 5, 5, 77.5, 16.5).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Know Payment"
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 12). _
        ParagraphFormat
        .FirstLineIndent = 0
        .Alignment = msoAlignLeft
    End With
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 12).Font
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Name = "+mn-lt"
    End With
    Selection.OnAction = "PERSONAL.XLSB!Module2.repaymentAmount"
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
    
Sheets("Loan_Schedule").Shapes.AddShape(msoShapeRectangle, 370, 6, 75, 26.5).Select
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
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Track  Payment"
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 14). _
        ParagraphFormat
        .FirstLineIndent = 0
        .Alignment = msoAlignLeft
    End With
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 14).Font
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 9
        .Name = "+mn-lt"
    End With
    Selection.OnAction = "PERSONAL.XLSB!Module2.trackRepay"
    
Sheets("Loan_Schedule").Shapes.AddShape(msoShapeRectangle, 852.5, 33, 43.5, 15).Select
    Selection.OnAction = "PERSONAL.XLSB!nomRateConversionLoan"
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
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 9
        .Name = "+mn-lt"
    End With
    
Sheets("Loan_Schedule").Shapes.AddShape(msoShapeRectangle, 850, 75.5, 43.5, 13).Select
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
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 9
        .Name = "+mn-lt"
    End With
    Selection.OnAction = "PERSONAL.XLSB!effRateConversionLoan"
    Selection.Formula = ""

Sheets("Loan_Schedule").Range("B1").Value = "Input"
Sheets("Loan_Schedule").Range("A2").Value = "Loan Type"
Sheets("Loan_Schedule").Range("A3").Value = "Amount Borrow"
Sheets("Loan_Schedule").Range("A4").Value = "Repayment Timeframe"
Sheets("Loan_Schedule").Range("A5").Value = "Time Horizon"
Sheets("Loan_Schedule").Range("A7").Value = "Repayment Deacription"
Sheets("Loan_Schedule").Range("A8").Value = "Interest Rate"
Sheets("Loan_Schedule").Range("C1").Value = "If Amortized Partially"
Sheets("Loan_Schedule").Range("C6").Value = "By"
Sheets("Loan_Schedule").Range("F1").Value = "Amount"
Sheets("Loan_Schedule").Range("G1").Value = "Period"
Sheets("Loan_Schedule").Range("A11").Value = "=""Period - "" " & "&R[-7]C[1]"
Sheets("Loan_Schedule").Range("B11").Value = "Outstanding Balance-Beginning"
Sheets("Loan_Schedule").Range("C11").Value = "Repayment Amount"
Sheets("Loan_Schedule").Range("D11").Value = "Interest Payment"
Sheets("Loan_Schedule").Range("E11").Value = "Capital Repay"
Sheets("Loan_Schedule").Range("F11").Value = "Outstanding Balance-End"
Sheets("Loan_Schedule").Range("I1").Value = "Rate Conversion calculator"
Sheets("Loan_Schedule").Range("I2").Value = "Nominal Interest"
Sheets("Loan_Schedule").Range("I3").Value = "=""Effective Interest Rate - ""&R[1]C[-7]"
Sheets("Loan_Schedule").Range("I5").Value = "Annual Effective Interest"
Sheets("Loan_Schedule").Range("I6").Value = "=""Effective Interest Rate - ""&R[-2]C[-7]"

Sheets("Loan_Schedule").Range("D10").Select
MsgBox "This calculator simplifies managing cashflows, making financial tasks quicker and easier." & vbCrLf & "Developer: Sekou M. Kamara", vbOKOnly + vbInformation, "Cashflow"
End Sub


