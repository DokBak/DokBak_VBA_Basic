Attribute VB_Name = "C05_Set_Conditional_Class"
Option Explicit


Function Set_Conditional_Sheet(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    
    ws.Name = SheetName
    
    Worksheets(SheetName).Activate
    
    For Each cell In Range("A2:I10")
        cell.Value = cell.Row() - 1
    Next cell
    
    For Each cell In Range("A13:A21")
        cell.Value = cell.Row() - 12
    Next cell
    For Each cell In Range("B13:E21")
        cell.Value = cell.Row() - 12 & " TestData " & cell.Row()
    Next cell
    
    For Each cell In Range("A24:A32")
        cell.Value = cell.Row() - 23
    Next cell
    For Each cell In Range("B24:B32")
        cell.Value = "TestData " & cell.Row()
    Next cell
    
    ws.Range("A2:I40").EntireColumn.AutoFit
    
End Function

Function Set_xlCellValue(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("A1").Value = "xlCellValue"
    
    ws.Range("B1").Value = "Cell > 5"
    ws.Range("B2:B10").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="5"
    ws.Range("B2:B10").FormatConditions(1).Interior.Color = vbRed
    
    ws.Range("C1").Value = "Cell >= 5"
    ws.Range("C2:C10").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="5"
    ws.Range("C2:C10").FormatConditions(1).Interior.Color = vbGreen
    
    ws.Range("D1").Value = "Cell = 5"
    ws.Range("D2:D10").FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="5"
    ws.Range("D2:D10").FormatConditions(1).Interior.Color = vbBlue
    
    ws.Range("E1").Value = "Cell <= 5"
    ws.Range("E2:E10").FormatConditions.Add Type:=xlCellValue, Operator:=xlLessEqual, Formula1:="5"
    ws.Range("E2:E10").FormatConditions(1).Interior.Color = vbYellow
    
    ws.Range("F1").Value = "Cell < 5"
    ws.Range("F2:F10").FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="5"
    ws.Range("F2:F10").FormatConditions(1).Interior.Color = vbMagenta
    
    ws.Range("G1").Value = "Cell > 5"
    ws.Range("G2:G10").FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="5"
    ws.Range("G2:G10").FormatConditions(1).Interior.Color = vbCyan
    
    ws.Range("H1").Value = "Cell >= 3, Cell <= 7"
    ws.Range("H2:H10").FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, Formula1:="3", Formula2:="7"
    ws.Range("H2:H10").FormatConditions(1).Interior.Color = RGB(200, 100, 50)
    
    ws.Range("I1").Value = "Cell <= 3, Cell >= 7"
    ws.Range("I2:I10").FormatConditions.Add Type:=xlCellValue, Operator:=xlNotBetween, Formula1:="3", Formula2:="7"
    ws.Range("I2:I10").FormatConditions(1).Interior.Color = RGB(100, 50, 200)
    
    ws.Range("A2:I10").EntireColumn.AutoFit

End Function

Function Set_xlExpression(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("A12").Value = "xlTextString"
    
    ws.Range("B12").Value = "xlContains 5"
    ws.Range("B13:B21").FormatConditions.Add Type:=xlTextString, TextOperator:=xlContains, String:="5"
    ws.Range("B13:B21").FormatConditions(1).Interior.Color = vbRed
    
    ws.Range("C12").Value = "xlDoesNotContain 5"
    ws.Range("C13:C21").FormatConditions.Add Type:=xlTextString, TextOperator:=xlDoesNotContain, String:="5"
    ws.Range("C13:C21").FormatConditions(1).Interior.Color = vbGreen
    
    ws.Range("D12").Value = "xlBeginsWith 5"
    ws.Range("D13:D21").FormatConditions.Add Type:=xlTextString, TextOperator:=xlBeginsWith, String:="5"
    ws.Range("D13:D21").FormatConditions(1).Interior.Color = vbBlue

    ws.Range("E12").Value = "xlEndsWith 5"
    ws.Range("E13:E21").FormatConditions.Add Type:=xlTextString, TextOperator:=xlEndsWith, String:="5"
    ws.Range("E13:E21").FormatConditions(1).Interior.Color = vbYellow

    ws.Range("A13:E21").EntireColumn.AutoFit

End Function


