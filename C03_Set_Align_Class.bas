Attribute VB_Name = "C03_Set_Align_Class"
Option Explicit

Function Set_Align_Sheet(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    
    ws.Name = SheetName
    
    Worksheets(SheetName).Activate
    
    ws.Range("A1") = "Default"
    For Each cell In Range("A2:A10")
        cell.Value = "Default" & cell.Row()
    Next cell
    
End Function

Function Set_Horizontal_Alignment(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("B1") = "HorizontalAlignment"
    For Each cell In Range("B2:B10")
        If cell.Row() < 3 Then
            cell.Value = "Horizontal Alignment " & "xlHAlignLeft"
        ElseIf cell.Row() < 4 Then
            cell.Value = "Horizontal Alignment " & "xlHAlignCenter"
        ElseIf cell.Row() < 5 Then
            cell.Value = "Horizontal Alignment " & "xlHAlignRight"
        ElseIf cell.Row() < 6 Then
            cell.Value = "Horizontal Alignment " & "xlHAlignFill"
        ElseIf cell.Row() < 7 Then
            cell.Value = "Horizontal Alignment " & "xlHAlignJustify"
        ElseIf cell.Row() < 8 Then
            cell.Value = "Horizontal Alignment " & "xlHAlignCenterAcrossSelection"
        ElseIf cell.Row() < 9 Then
            cell.Value = "Horizontal Alignment " & "xlHAlignDistributed"
        Else
            cell.Value = "Horizontal Alignment " & "xlHAlignGeneral"
        End If
    Next cell
    
    ws.Range("B2:B10").HorizontalAlignment = xlHAlignLeft
    ws.Range("B3:B10").HorizontalAlignment = xlHAlignCenter
    ws.Range("B4:B10").HorizontalAlignment = xlHAlignRight
    ws.Range("B5:B10").HorizontalAlignment = xlHAlignFill
    ws.Range("B6:B10").HorizontalAlignment = xlHAlignJustify
    ws.Range("B7:B10").HorizontalAlignment = xlHAlignCenterAcrossSelection
    ws.Range("B8:B10").HorizontalAlignment = xlHAlignDistributed
    ws.Range("B9:B10").HorizontalAlignment = xlHAlignGeneral
    
    ws.Range("B1:D10").ColumnWidth = 50
    ws.Range("B1:D10").EntireColumn.AutoFit
    
End Function

Function Set_Vertical_Alignment(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("C1") = "VerticalAlignment"
    For Each cell In Range("C2:C10")
        If cell.Row() < 3 Then
            cell.Value = "Vertical Alignment " & "xlVAlignTop"
        ElseIf cell.Row() < 4 Then
            cell.Value = "Vertical Alignment " & "xlCenter"
        ElseIf cell.Row() < 5 Then
            cell.Value = "Vertical Alignment " & "xlVAlignBottom"
        ElseIf cell.Row() < 6 Then
            cell.Value = "Vertical Alignment " & vbNewLine & "xlVAlignJustify"
        ElseIf cell.Row() < 7 Then
            cell.Value = "Vertical Alignment " & vbNewLine & "xlVAlignDistributed"
        Else
            cell.Value = "Vertical Alignment " & vbNewLine & "xlGeneral"
        End If
    Next cell
    ws.Range("C2:C10").VerticalAlignment = xlVAlignTop
    ws.Range("C3:C10").VerticalAlignment = xlCenter
    ws.Range("C4:C10").VerticalAlignment = xlVAlignBottom
    ws.Range("C5:C10").VerticalAlignment = xlVAlignJustify
    ws.Range("C6:C10").VerticalAlignment = xlVAlignDistributed
    ws.Range("C7:C10").VerticalAlignment = xlGeneral
    
    ws.Range("C1:C10").RowHeight = 70
    ws.Range("C1:C10").EntireColumn.AutoFit
    
End Function

Function Set_IndentLevel(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("D1") = "IndentLevel"
    For Each cell In Range("D2:D10")
        If cell.Row() < 3 Then
            cell.Value = "IndentLevel " & "1 " & "xlHAlignLeft"
        ElseIf cell.Row() < 4 Then
            cell.Value = "IndentLevel " & "2 " & "xlHAlignLeft"
        ElseIf cell.Row() < 5 Then
            cell.Value = "IndentLevel " & "0 " & "xlHAlignRight"
        ElseIf cell.Row() < 6 Then
            cell.Value = "IndentLevel " & "1 " & "xlHAlignRight"
        ElseIf cell.Row() < 7 Then
            cell.Value = "IndentLevel " & "2 " & "xlHAlignRight"
        ElseIf cell.Row() < 8 Then
            cell.Value = "IndentLevel " & "0 " & "xlHAlignRight"
        ElseIf cell.Row() < 9 Then
            cell.Value = "IndentLevel " & "1 " & "xlGeneral"
        Else
            cell.Value = "IndentLevel " & "2 " & "xlGeneral"
        End If
    Next cell
    
    ws.Range("D2:D10").HorizontalAlignment = xlHAlignLeft
    ws.Range("D2:D10").IndentLevel = 1
    ws.Range("D3:D10").IndentLevel = 2
    ws.Range("D4:D10").IndentLevel = 0
    ws.Range("D5:D10").HorizontalAlignment = xlHAlignRight
    ws.Range("D5:D10").IndentLevel = 1
    ws.Range("D6:D10").IndentLevel = 2
    ws.Range("D7:D10").IndentLevel = 0
    ws.Range("D8:D10").HorizontalAlignment = xlGeneral
    ws.Range("D9:D10").IndentLevel = 1
    ws.Range("D10:D10").IndentLevel = 2
    
    ws.Range("D1:D10").EntireColumn.AutoFit
    
End Function

Function Set_Orientation(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("E1") = "Orientation"
    For Each cell In Range("E2:E10")
        If cell.Row() < 3 Then
            cell.Value = "Orientation " & "xlVertical"
        ElseIf cell.Row() < 4 Then
            cell.Value = "Orientation " & "xlHorizontal"
        ElseIf cell.Row() < 5 Then
            cell.Value = "Orientation " & "90"
        ElseIf cell.Row() < 6 Then
            cell.Value = "Orientation " & "-90"
        ElseIf cell.Row() < 7 Then
            cell.Value = "Orientation " & "45"
        ElseIf cell.Row() < 8 Then
            cell.Value = "Orientation " & "-45"
        ElseIf cell.Row() < 9 Then
            cell.Value = "Orientation " & "xlGeneral"
        Else
            cell.Value = "Orientation " & "xlGeneral"
        End If
    Next cell
    
    ws.Range("E2:E10").Orientation = xlVertical
    ws.Range("E3:E10").Orientation = xlHorizontal
    ws.Range("E4:E10").Orientation = 90
    ws.Range("E5:E10").Orientation = -90
    ws.Range("E6:E10").Orientation = 45
    ws.Range("E7:E10").Orientation = -45
    ws.Range("E8:E10").Orientation = xlGeneral
    
    ws.Range("E1:E10").EntireColumn.AutoFit
    
End Function

Function Set_WrapText(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("F1") = "WrapText"
    For Each cell In Range("F2:F10")
        If cell.Row() < 7 Then
            cell.Value = "WrapText " & "True"
        ElseIf cell.Row() < 10 Then
            cell.Value = "WrapText " & "False"
        Else
            cell.Value = "WrapText " & "xlGeneral"
        End If
    Next cell
    
    ws.Range("F2:F10").WrapText = True
    ws.Range("F6:F10").WrapText = False
    ws.Range("F10:F10").WrapText = xlGeneral
    
    ws.Range("F1:F10").EntireColumn.AutoFit
    ws.Range("F1:F10").ColumnWidth = 8
    
End Function

Function Set_Merge(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("G1:H1") = "Merge"
    For Each cell In Range("G2:G10")
        If cell.Row() < 5 Then
            cell.Value = "Merge"
        ElseIf cell.Row() < 8 Then
            cell.Value = "Merge & xlCenter"
        Else
            cell.Value = "Merge & UnMerge"
        End If
    Next cell
    
    ws.Range("G2:H4").Merge
    ws.Range("G5:H7").Merge
    ws.Range("G5:H7").HorizontalAlignment = xlCenter
    ws.Range("G8:G10").Merge
    ws.Range("G10:H10").UnMerge
    
    ws.Range("G1:H10").EntireColumn.AutoFit
    
End Function

Function Set_AutoFit(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("I1") = "AutoFit"
    ws.Range("J1") = "AutoFit"
    For Each cell In Range("I2:J10")
        If cell.Row() < 6 Then
            cell.Value = "Columns.AutoFit"
        Else
            cell.Value = "Rows.AutoFit"
        End If
    Next cell
    
    ws.Range("I2:I10").Columns.AutoFit
    ws.Range("J2:J10").Rows.AutoFit

End Function
