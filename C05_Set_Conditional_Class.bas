Attribute VB_Name = "C05_Set_Conditional_Class"
Option Explicit

Function Set_Conditional_Sheet(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    
    ws.Name = SheetName
    
    Worksheets(SheetName).Activate
    
    'Set_xlCellValue
    For Each cell In Range("A2:I11")
        cell.Value = cell.Row() - 1
    Next cell
    
    'Set_xlTextString
    For Each cell In Range("A14:A23")
        cell.Value = cell.Row() - 13
    Next cell
    For Each cell In Range("B14:E23")
        cell.Value = cell.Row() - 13 & " TestData " & cell.Row()
    Next cell
    
    'Set_xlTimePeriod
    For Each cell In Range("A26:A36")
        cell.Value = cell.Row() - 25
    Next cell
    For Each cell In Range("B26:B36")
        If cell.Row() - 25 = 1 Then
            cell.Value = "Today"
        ElseIf cell.Row() - 25 = 2 Then
            cell.Value = "1 day ago"
        ElseIf cell.Row() - 25 = 3 Then
            cell.Value = "1 day later"
        ElseIf cell.Row() - 25 = 4 Then
            cell.Value = "3 days ago"
        ElseIf cell.Row() - 25 = 5 Then
            cell.Value = "3 days later"
        ElseIf cell.Row() - 25 = 6 Then
            cell.Value = "7 days ago"
        ElseIf cell.Row() - 25 = 7 Then
            cell.Value = "7 days later"
        ElseIf cell.Row() - 25 = 8 Then
            cell.Value = "1 month ago"
        ElseIf cell.Row() - 25 = 9 Then
            cell.Value = "1 month later"
        ElseIf cell.Row() - 25 = 10 Then
            cell.Value = "1 year ago"
        ElseIf cell.Row() - 25 = 11 Then
            cell.Value = "1 year later"
        Else
            cell.Value = "Today"
        End If
    Next cell
    For Each cell In Range("C26:L36")
        If cell.Row() - 25 = 1 Then
            cell.Value = Date
        ElseIf cell.Row() - 25 = 2 Then
            cell.Value = Date - 1
        ElseIf cell.Row() - 25 = 3 Then
            cell.Value = Date + 1
        ElseIf cell.Row() - 25 = 4 Then
            cell.Value = Date - 3
        ElseIf cell.Row() - 25 = 5 Then
            cell.Value = Date + 3
        ElseIf cell.Row() - 25 = 6 Then
            cell.Value = Date + 7
        ElseIf cell.Row() - 25 = 7 Then
            cell.Value = Date - 7
        ElseIf cell.Row() - 25 = 8 Then
            cell.Value = DateAdd("m", -1, Date)
        ElseIf cell.Row() - 25 = 9 Then
            cell.Value = DateAdd("m", 1, Date)
        ElseIf cell.Row() - 25 = 10 Then
            cell.Value = DateAdd("yyyy", -1, Date)
        ElseIf cell.Row() - 25 = 11 Then
            cell.Value = DateAdd("yyyy", 1, Date)
        Else
            cell.Value = Date
        End If
    Next cell
    
    'Set_xlBlanksCondition
    For Each cell In Range("A39:B43")
        cell.Value = cell.Row() - 38
    Next cell
    
    'Set_xlNoErrorsCondition
    For Each cell In Range("A46:A55")
        If cell.Row() - 45 <= 6 Then
            cell.Value = cell.Row() - 45
        Else
            cell.Value = cell.Row() - 47
        End If
    Next cell
    For Each cell In Range("B46:C53")
        If cell.Row() - 45 = 1 Then
            cell.Value = "=2/0" '#DIV/0!
        ElseIf cell.Row() - 45 = 2 Then
            cell.Value = "=d/2" '#NAME?
        ElseIf cell.Row() - 45 = 3 Then
            cell.Value = "=VLOOKUP(,,,)"  '#N/A
        ElseIf cell.Row() - 45 = 4 Then
            cell.Value = "=VAR(""&dd&"")" '#VALUE!
        ElseIf cell.Row() - 45 = 5 Then
            cell.Value = "=A51" '#REF!
            Range("A51").Delete Shift:=xlUp
        ElseIf cell.Row() - 45 = 6 Then
            cell.Value = "=9.9*10^308" '#NUM!
        ElseIf cell.Row() - 45 = 7 Then
            cell.Value = "=SUM(A46:A47 A49:A50)" '#NULL!
        Else
            cell.Value = "=Today()"
        End If
    Next cell
    
    'Set_xlTop10Top
    For Each cell In Range("A56:E75")
        cell.Value = cell.Row() - 55
    Next cell
    
    'Set_AddAboveAverage
    For Each cell In Range("A78:G88")
        cell.Value = cell.Row() - 77
    Next cell
    
    'Set_AddUniqueValues
    For Each cell In Range("A91:A100")
        cell.Value = cell.Row() - 90
    Next cell
    For Each cell In Range("B91:C100")
        If cell.Row() - 90 = 1 Then
            cell.Value = "Data1"
        ElseIf cell.Row() - 90 < 5 Then
            cell.Value = "Data2"
        ElseIf cell.Row() - 90 = 5 Then
            cell.Value = "Data3"
        ElseIf cell.Row() - 90 < 10 Then
            cell.Value = "Data4"
        ElseIf cell.Row() - 90 = 10 Then
            cell.Value = "DATA1"
        Else
            cell.Value = "Data6"
        End If
    Next cell
    
    'Set_xlExpression
    For Each cell In Range("A103:A112")
        cell.Value = cell.Row() - 102
    Next cell
    For Each cell In Range("B103:B112")
        cell.Value = cell.Row() * 4.5
    Next cell
    
    ws.Range("A2:I200").EntireColumn.AutoFit
    
End Function

Function Set_xlCellValue(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("A1").Value = "xlCellValue"
    
    ws.Range("B1").Value = "Cell > 5"
    ws.Range("B2:B11").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="5"
    ws.Range("B2:B11").FormatConditions(1).Interior.Color = vbRed
    
    ws.Range("C1").Value = "Cell >= 5"
    ws.Range("C2:C11").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="5"
    ws.Range("C2:C11").FormatConditions(1).Interior.Color = vbGreen
    
    ws.Range("D1").Value = "Cell = 5"
    ws.Range("D2:D11").FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="5"
    ws.Range("D2:D11").FormatConditions(1).Interior.Color = vbBlue
    
    ws.Range("E1").Value = "Cell <= 5"
    ws.Range("E2:E11").FormatConditions.Add Type:=xlCellValue, Operator:=xlLessEqual, Formula1:="5"
    ws.Range("E2:E11").FormatConditions(1).Interior.Color = vbYellow
    
    ws.Range("F1").Value = "Cell < 5"
    ws.Range("F2:F11").FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="5"
    ws.Range("F2:F11").FormatConditions(1).Interior.Color = vbMagenta
    
    ws.Range("G1").Value = "Cell <> 5"
    ws.Range("G2:G11").FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="5"
    ws.Range("G2:G11").FormatConditions(1).Interior.Color = vbCyan
    
    ws.Range("H1").Value = "Cell >= 3, Cell <= 7"
    ws.Range("H2:H11").FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, Formula1:="3", Formula2:="7"
    ws.Range("H2:H11").FormatConditions(1).Interior.Color = RGB(200, 100, 50)
    
    ws.Range("I1").Value = "Cell < 3, Cell > 7"
    ws.Range("I2:I11").FormatConditions.Add Type:=xlCellValue, Operator:=xlNotBetween, Formula1:="3", Formula2:="7"
    ws.Range("I2:I11").FormatConditions(1).Interior.Color = RGB(100, 50, 200)
    
    ws.Range("A1:I11").EntireColumn.AutoFit

End Function

Function Set_xlTextString(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("A13").Value = "xlTextString"
    
    ws.Range("B13").Value = "xlContains 5"
    ws.Range("B14:B23").FormatConditions.Add Type:=xlTextString, TextOperator:=xlContains, String:="5"
    ws.Range("B14:B23").FormatConditions(1).Interior.Color = vbRed
    
    ws.Range("C13").Value = "xlDoesNotContain 5"
    ws.Range("C14:C23").FormatConditions.Add Type:=xlTextString, TextOperator:=xlDoesNotContain, String:="5"
    ws.Range("C14:C23").FormatConditions(1).Interior.Color = vbGreen
    
    ws.Range("D13").Value = "xlBeginsWith 5"
    ws.Range("D14:D23").FormatConditions.Add Type:=xlTextString, TextOperator:=xlBeginsWith, String:="5"
    ws.Range("D14:D23").FormatConditions(1).Interior.Color = vbBlue

    ws.Range("E13").Value = "xlEndsWith 5"
    ws.Range("E14:E23").FormatConditions.Add Type:=xlTextString, TextOperator:=xlEndsWith, String:="5"
    ws.Range("E14:E23").FormatConditions(1).Interior.Color = vbYellow

    ws.Range("A13:E23").EntireColumn.AutoFit

End Function

Function Set_xlTimePeriod(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("A25").Value = "xlTimePeriod"
    ws.Range("B25").Value = "Date"
    
    ws.Range("C25").Value = "xlToday"
    ws.Range("C26:C36").FormatConditions.Add Type:=xlTimePeriod, DateOperator:=xlToday
    ws.Range("C26:C36").FormatConditions(1).Interior.Color = vbRed

    ws.Range("D25").Value = "xlYesterday"
    ws.Range("D26:D36").FormatConditions.Add Type:=xlTimePeriod, DateOperator:=xlYesterday
    ws.Range("D26:D36").FormatConditions(1).Interior.Color = vbGreen

    ws.Range("E25").Value = "xlTomorrow"
    ws.Range("E26:E36").FormatConditions.Add Type:=xlTimePeriod, DateOperator:=xlTomorrow
    ws.Range("E26:E36").FormatConditions(1).Interior.Color = vbBlue

    ws.Range("F25").Value = "xlLast7Days"
    ws.Range("F26:F36").FormatConditions.Add Type:=xlTimePeriod, DateOperator:=xlLast7Days
    ws.Range("F26:F36").FormatConditions(1).Interior.Color = vbYellow
    
    ws.Range("G25").Value = "xlThisWeek"
    ws.Range("G26:G36").FormatConditions.Add Type:=xlTimePeriod, DateOperator:=xlThisWeek
    ws.Range("G26:G36").FormatConditions(1).Interior.Color = vbMagenta
    
    ws.Range("H25").Value = "xlLastWeek"
    ws.Range("H26:H36").FormatConditions.Add Type:=xlTimePeriod, DateOperator:=xlLastWeek
    ws.Range("H26:H36").FormatConditions(1).Interior.Color = vbCyan
    
    ws.Range("I25").Value = "xlNextWeek"
    ws.Range("I26:I36").FormatConditions.Add Type:=xlTimePeriod, DateOperator:=xlNextWeek
    ws.Range("I26:I36").FormatConditions(1).Interior.Color = RGB(200, 100, 50)
    
    ws.Range("J25").Value = "xlThisMonth"
    ws.Range("J26:J36").FormatConditions.Add Type:=xlTimePeriod, DateOperator:=xlThisMonth
    ws.Range("J26:J36").FormatConditions(1).Interior.Color = RGB(100, 50, 200)
    
    ws.Range("K25").Value = "xlLastMonth"
    ws.Range("K26:K36").FormatConditions.Add Type:=xlTimePeriod, DateOperator:=xlLastMonth
    ws.Range("K26:K36").FormatConditions(1).Interior.Color = RGB(50, 50, 50)
    
    ws.Range("L25").Value = "xlNextMonth"
    ws.Range("L26:L36").FormatConditions.Add Type:=xlTimePeriod, DateOperator:=xlNextMonth
    ws.Range("L26:L36").FormatConditions(1).Interior.Color = RGB(100, 100, 100)
    
    ws.Range("J25").Value = "xlThisMonth"
    ws.Range("J26:J36").FormatConditions.Add Type:=xlTimePeriod, DateOperator:=xlYear
    ws.Range("J26:J36").FormatConditions(1).Interior.Color = RGB(150, 150, 150)
    
    ws.Range("A25:O36").EntireColumn.AutoFit

End Function

Function Set_xlBlanksCondition(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("A38").Value = "xlBlanksCondition"
    
    ws.Range("B38").Value = "xlNoBlanksCondition"
    ws.Range("B39:B43").FormatConditions.Add Type:=xlNoBlanksCondition
    ws.Range("B39:B43").FormatConditions(1).Interior.Color = vbRed
    
    ws.Range("C38").Value = "xlBlanksCondition"
    ws.Range("C39:C43").FormatConditions.Add Type:=xlBlanksCondition
    ws.Range("C39:C43").FormatConditions(1).Interior.Color = vbGreen
    
    ws.Range("A38:C43").EntireColumn.AutoFit

End Function

Function Set_xlNoErrorsCondition(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("A45").Value = "xlErrorsCondition"
    
    ws.Range("B45").Value = "xlNoErrorsCondition"
    ws.Range("B46:B53").FormatConditions.Add Type:=xlNoErrorsCondition
    ws.Range("B46:B53").FormatConditions(1).Interior.Color = vbRed
    
    ws.Range("C45").Value = "xlErrorsCondition"
    ws.Range("C46:C53").FormatConditions.Add Type:=xlErrorsCondition
    ws.Range("C46:C53").FormatConditions(1).Interior.Color = vbGreen
    
    ws.Range("A45:C53").EntireColumn.AutoFit

End Function

Function Set_xlTop10Top(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("A55").Value = "xlTop10Top"
    
    ws.Range("B55").Value = "xlTop10Top_Count"
    ws.Range("B56:B75").FormatConditions.AddTop10
    With Range("B56:B75").FormatConditions(ws.Range("B56:B75").FormatConditions.Count)
        .TopBottom = xlTop10Top
        .Rank = 15
        .Percent = False
    End With
    ws.Range("B56:B75").FormatConditions(ws.Range("B56:B75").FormatConditions.Count).Interior.Color = vbRed

    ws.Range("C55").Value = "xlTop10Top_Percent"
    ws.Range("C56:C75").FormatConditions.AddTop10
    With Range("C56:C75").FormatConditions(ws.Range("C56:C75").FormatConditions.Count)
        .TopBottom = xlTop10Top
        .Rank = 15
        .Percent = True
    End With
    ws.Range("C56:C75").FormatConditions(ws.Range("C56:C75").FormatConditions.Count).Interior.Color = vbGreen
    
    ws.Range("D55").Value = "xlTop10Bottom_Count"
    ws.Range("D56:D75").FormatConditions.AddTop10
    With Range("D56:D75").FormatConditions(ws.Range("D56:D75").FormatConditions.Count)
        .TopBottom = xlTop10Bottom
        .Rank = 15
        .Percent = False
    End With
    ws.Range("D56:D75").FormatConditions(ws.Range("D56:D75").FormatConditions.Count).Interior.Color = vbBlue
    
    ws.Range("E55").Value = "xlTop10Bottom_Percent"
    ws.Range("E56:E75").FormatConditions.AddTop10
    With Range("E56:E75").FormatConditions(ws.Range("E56:E75").FormatConditions.Count)
        .TopBottom = xlTop10Bottom
        .Rank = 15
        .Percent = True
    End With
    ws.Range("E56:E75").FormatConditions(ws.Range("E56:E75").FormatConditions.Count).Interior.Color = vbYellow
    
    ws.Range("B56:E75").EntireColumn.AutoFit
    
End Function

Function Set_AddAboveAverage(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("A77").Value = "AddAboveAverage"
    
    ws.Range("B77").Value = "xlAboveAverage"
    With ws.Range("B78:B88").FormatConditions.AddAboveAverage
      .AboveBelow = xlAboveAverage
      With .Interior
        .Color = vbRed
      End With
    End With
    
    ws.Range("C77").Value = "xlEqualAboveAverage"
    With ws.Range("C78:C88").FormatConditions.AddAboveAverage
      .AboveBelow = xlEqualAboveAverage
      With .Interior
        .Color = vbGreen
      End With
    End With
    
    ws.Range("D77").Value = "xlBelowAverage"
    With ws.Range("D78:D88").FormatConditions.AddAboveAverage
      .AboveBelow = xlBelowAverage
      With .Interior
        .Color = vbBlue
      End With
    End With
    
    ws.Range("E77").Value = "xlEqualBelowAverage"
    With ws.Range("E78:E88").FormatConditions.AddAboveAverage
      .AboveBelow = xlEqualBelowAverage
      With .Interior
        .Color = vbYellow
      End With
    End With
    
    ws.Range("F77").Value = "xlAboveStdDev"
    With ws.Range("F78:F88").FormatConditions.AddAboveAverage
      .AboveBelow = xlAboveStdDev
      With .Interior
        .Color = vbMagenta
      End With
    End With
    
    ws.Range("G77").Value = "xlBelowStdDev"
    With ws.Range("G78:G88").FormatConditions.AddAboveAverage
      .AboveBelow = xlBelowStdDev
      With .Interior
        .Color = vbCyan
      End With
    End With
    
    ws.Range("A77:G88").EntireColumn.AutoFit

End Function

Function Set_AddUniqueValues(SheetName As String)
    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("A90").Value = "AddUniqueValues"
    
    ws.Range("B90").Value = "xlUnique"
    With Range("B91:B100").FormatConditions.AddUniqueValues
       .DupeUnique = xlUnique
       .Interior.Color = vbRed
    End With
    
    ws.Range("C90").Value = "xlDuplicate"
    With Range("C91:C100").FormatConditions.AddUniqueValues
       .DupeUnique = xlDuplicate
       .Interior.Color = vbGreen
    End With
     
    ws.Range("A90:C100").EntireColumn.AutoFit
    
End Function

Function Set_xlExpression(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("A102").Value = "xlExpression"
    
    ws.Range("B102").Value = "xlExpression"
    Range("B103:B112").FormatConditions.Add Type:=xlExpression, Formula1:="=LEN(TRIM(B103:B112))>4"
    Range("B103:B112").FormatConditions(1).Interior.Color = vbRed
    
    ws.Range("A102:B112").EntireColumn.AutoFit
    
End Function

