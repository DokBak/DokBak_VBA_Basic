Attribute VB_Name = "C04_Set_Format_Class"
Option Explicit

Function Set_Format_Sheet(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    
    ws.Name = SheetName
    
    Worksheets(SheetName).Activate
    
    ws.Range("A1") = "Format Type"
    For Each cell In Range("A2:A25")
        If cell.Row() < 3 Then
            cell.Value = "General"
        ElseIf cell.Row() < 4 Then
            cell.Value = "Float"
        ElseIf cell.Row() < 5 Then
            cell.Value = "Float"
        ElseIf cell.Row() < 6 Then
            cell.Value = "Float"
        ElseIf cell.Row() < 7 Then
            cell.Value = "Float"
        ElseIf cell.Row() < 8 Then
            cell.Value = "Float"
        ElseIf cell.Row() < 9 Then
            cell.Value = "Fraction"
        ElseIf cell.Row() < 10 Then
            cell.Value = "Float"
        ElseIf cell.Row() < 11 Then
            cell.Value = "Float"
        ElseIf cell.Row() < 12 Then
            cell.Value = "Percentage"
        ElseIf cell.Row() < 13 Then
            cell.Value = "Currency"
        ElseIf cell.Row() < 14 Then
            cell.Value = "Currency"
        ElseIf cell.Row() < 15 Then
            cell.Value = "Currency"
        ElseIf cell.Row() < 16 Then
            cell.Value = "Accounting"
        ElseIf cell.Row() < 17 Then
            cell.Value = "Date"
        ElseIf cell.Row() < 18 Then
            cell.Value = "Date"
        ElseIf cell.Row() < 19 Then
            cell.Value = "Date"
        ElseIf cell.Row() < 20 Then
            cell.Value = "Date"
        ElseIf cell.Row() < 21 Then
            cell.Value = "Time"
        ElseIf cell.Row() < 22 Then
            cell.Value = "Time"
        ElseIf cell.Row() < 23 Then
            cell.Value = "Time"
        ElseIf cell.Row() < 24 Then
            cell.Value = "Time"
        ElseIf cell.Row() < 25 Then
            cell.Value = "String"
        Else
            cell.Value = "Custom Format"
        End If
    Next cell
    
    ws.Range("A1:A25").EntireColumn.AutoFit
    
End Function
Function Set_NumberFormat_Format(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("B1") = "NumberFormat"
    
    ws.Range("B2:B25").NumberFormat = "@"

    For Each cell In Range("B2:B25")
        If cell.Row() < 3 Then
            cell.Value = "General"
        ElseIf cell.Row() < 4 Then
            cell.Value = "0000.0000"
        ElseIf cell.Row() < 5 Then
            cell.Value = "###0.0###"
        ElseIf cell.Row() < 6 Then
            cell.Value = "#,##0.##;[Red]-#,##0.0##"
        ElseIf cell.Row() < 7 Then
            cell.Value = "##,#0.##;[Red]-##,#0.0##"
        ElseIf cell.Row() < 8 Then
            cell.Value = "##,#0.##;[Blue]-##,#0.0##"
        ElseIf cell.Row() < 9 Then
            cell.Value = "# ?/?"
        ElseIf cell.Row() < 10 Then
            cell.Value = "0.00E+00"
        ElseIf cell.Row() < 11 Then
            cell.Value = "0E+00"
        ElseIf cell.Row() < 12 Then
            cell.Value = "##0.00%"
        ElseIf cell.Row() < 13 Then
            cell.Value = "$#,##0.00"
        ElseIf cell.Row() < 14 Then
            cell.Value = "£Ü#,###0.00"
        ElseIf cell.Row() < 15 Then
            cell.Value = "¡Í#,###0.00"
        ElseIf cell.Row() < 16 Then
            cell.Value = "_-[$£Ü-412]* #,##0.00_ ;_-[$£Ü-412]* -#,##0.00 ;_-[$£Ü-412]* ""-""??_ ;_-@_-"
        ElseIf cell.Row() < 17 Then
            cell.Value = "YYYYMMDD"
        ElseIf cell.Row() < 18 Then
            cell.Value = "YYYY/MM/DD"
        ElseIf cell.Row() < 19 Then
            cell.Value = "YY/M/D"
        ElseIf cell.Row() < 20 Then
            cell.Value = "YY/M/D"
        ElseIf cell.Row() < 21 Then
            cell.Value = "hhmmss"
        ElseIf cell.Row() < 22 Then
            cell.Value = "hh:mm:ss"
        ElseIf cell.Row() < 23 Then
            cell.Value = "h:m:s"
        ElseIf cell.Row() < 24 Then
            cell.Value = "h:m:s"
        ElseIf cell.Row() < 25 Then
            cell.Value = "@"
        Else
            cell.Value = "[Red]@"
        End If
    Next cell
    
    ws.Range("B1:B25").EntireColumn.AutoFit

End Function
Function Set_Data_Format(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("C1") = "Data"
    
    ws.Range("C2:C25").NumberFormat = "@"
    
    For Each cell In Range("C2:C25")
        If cell.Row() < 3 Then
            cell.Value = 45
        ElseIf cell.Row() < 4 Then
            cell.Value = 345.673
        ElseIf cell.Row() < 5 Then
            cell.Value = 345.673
        ElseIf cell.Row() < 6 Then
            cell.Value = -345.673
        ElseIf cell.Row() < 7 Then
            cell.Value = -345.673
        ElseIf cell.Row() < 8 Then
            cell.Value = -345.673
        ElseIf cell.Row() < 9 Then
            cell.Value = 1 / 3
        ElseIf cell.Row() < 10 Then
            cell.Value = 1 / 3
        ElseIf cell.Row() < 11 Then
            cell.Value = 100
        ElseIf cell.Row() < 12 Then
            cell.Value = 100
        ElseIf cell.Row() < 13 Then
            cell.Value = 100000
        ElseIf cell.Row() < 14 Then
            cell.Value = 100000
        ElseIf cell.Row() < 15 Then
            cell.Value = 100000
        ElseIf cell.Row() < 16 Then
            cell.Value = 100000
        ElseIf cell.Row() < 17 Then
            cell.Value = "2023-04-01"
        ElseIf cell.Row() < 18 Then
            cell.Value = "2023/04/01"
        ElseIf cell.Row() < 19 Then
            cell.Value = "2023/04/01"
        ElseIf cell.Row() < 20 Then
            cell.Value = "2023/11/21"
        ElseIf cell.Row() < 21 Then
            cell.Value = "07:08:09"
        ElseIf cell.Row() < 22 Then
            cell.Value = "07:08:09"
        ElseIf cell.Row() < 23 Then
            cell.Value = "07:08:09"
        ElseIf cell.Row() < 24 Then
            cell.Value = "11:23:35"
        ElseIf cell.Row() < 25 Then
            cell.Value = "TEXT DATA"
        Else
            cell.Value = "Color Red DATA"
        End If
    Next cell
    
    ws.Range("C1:C25").EntireColumn.AutoFit

End Function

Function Set_Example_Format(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("D1") = "Example"

    For Each cell In Range("D2:D25")
        If cell.Row() < 3 Then
            cell.Value = 45
            cell.NumberFormat = "General"
        ElseIf cell.Row() < 4 Then
            cell.Value = 345.673
            cell.NumberFormat = "0000.0000"
        ElseIf cell.Row() < 5 Then
            cell.Value = 345.673
            cell.NumberFormat = "###0.0###"
        ElseIf cell.Row() < 6 Then
            cell.Value = -345.673
            cell.NumberFormat = "#,##0.##;[Red]-#,##0.0##"
        ElseIf cell.Row() < 7 Then
            cell.Value = -345.673
            cell.NumberFormat = "##,#0.##;[Red]-##,#0.0##"
        ElseIf cell.Row() < 8 Then
            cell.Value = -345.673
            cell.NumberFormat = "##,#0.##;[Blue]-##,#0.0##"
        ElseIf cell.Row() < 9 Then
            cell.Value = 1 / 3
            cell.NumberFormat = "# ?/?"
        ElseIf cell.Row() < 10 Then
            cell.Value = 1 / 3
            cell.NumberFormat = "0.00E+00"
        ElseIf cell.Row() < 11 Then
            cell.Value = 100
            cell.NumberFormat = "0E+00"
        ElseIf cell.Row() < 12 Then
            cell.Value = 100
            cell.NumberFormat = "##0.00%"
        ElseIf cell.Row() < 13 Then
            cell.Value = 100000
            cell.NumberFormat = "$#,##0.00"
        ElseIf cell.Row() < 14 Then
            cell.Value = 100000
            cell.NumberFormat = "£Ü#,###0.00"
        ElseIf cell.Row() < 15 Then
            cell.Value = 100000
            cell.NumberFormat = "¡Í#,###0.00"
        ElseIf cell.Row() < 16 Then
            cell.Value = 100000
            cell.NumberFormat = "_-[$£Ü-412]* #,##0.00_ ;_-[$£Ü-412]* -#,##0.00 ;_-[$£Ü-412]* ""-""??_ ;_-@_-"
        ElseIf cell.Row() < 17 Then
            cell.Value = "2023-04-01"
            cell.NumberFormat = "YYYYMMDD"
        ElseIf cell.Row() < 18 Then
            cell.Value = "2023/04/01"
            cell.NumberFormat = "YYYY/MM/DD"
        ElseIf cell.Row() < 19 Then
            cell.Value = "2023/04/01"
            cell.NumberFormat = "YY/M/D"
        ElseIf cell.Row() < 20 Then
            cell.Value = "2023/11/21"
            cell.NumberFormat = "YY/M/D"
        ElseIf cell.Row() < 21 Then
            cell.Value = "07:08:09"
            cell.NumberFormat = "hhmmss"
        ElseIf cell.Row() < 22 Then
            cell.Value = "07:08:09"
            cell.NumberFormat = "hh:mm:ss"
        ElseIf cell.Row() < 23 Then
            cell.Value = "07:08:09"
            cell.NumberFormat = "h:m:s"
        ElseIf cell.Row() < 24 Then
            cell.Value = "11:23:35"
            cell.NumberFormat = "h:m:s"
        ElseIf cell.Row() < 25 Then
            cell.Value = "TEXT DATA"
            cell.NumberFormat = "@"
        Else
            cell.Value = "Color Red DATA"
            cell.NumberFormat = "[Red]@"
        End If
    Next cell
    
    ws.Range("D1:D25").EntireColumn.AutoFit

End Function


