Attribute VB_Name = "C02_Set_Cell_Class"
Option Explicit

Function Set_Cell_Sheet(SheetName As String)
    
    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    
    ws.Name = SheetName
    
    Worksheets(SheetName).Activate
    
    ws.Range("A1") = "Default"
    For Each cell In Range("A2:A10")
        cell.Value = "Default " & cell.Row()
    Next cell
    ws.Range("A1:A10").EntireColumn.AutoFit
    
End Function
Function Set_Font_Name(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("B1") = "Font.Name"
    For Each cell In Range("B2:B10")
        If cell.Row() < 4 Then
            cell.Value = "Font.Name " & "Arial"
        ElseIf cell.Row() < 6 Then
            cell.Value = "Font.Name " & "Calibri"
        ElseIf cell.Row() < 8 Then
            cell.Value = "Font.Name " & "MS Mincho"
        Else
            cell.Value = "Font.Name " & "MS Gothic"
        End If
    Next cell
    ws.Range("B2:B10").Font.Name = "Arial"
    ws.Range("B4:B10").Font.Name = "Calibri"
    ws.Range("B6:B10").Font.Name = "MS Mincho"
    ws.Range("B8:B10").Font.Name = "MS Gothic"
    ws.Range("B1:B10").EntireColumn.AutoFit
    
End Function
Function Set_Font_Size(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("C1") = "Font.Size"
    For Each cell In Range("C2:C10")
        If cell.Row() < 4 Then
            cell.Value = "Font.Size " & 8
        ElseIf cell.Row() < 6 Then
            cell.Value = "Font.Size " & 10
        ElseIf cell.Row() < 8 Then
            cell.Value = "Font.Size " & 12
        Else
            cell.Value = "Font.Size " & 14.5
        End If
    Next cell
    ws.Range("C2:C10").Font.Size = 8
    ws.Range("C4:C10").Font.Size = 10
    ws.Range("C6:C10").Font.Size = 12
    ws.Range("C8:C10").Font.Size = 14.5
    ws.Range("C1:C10").EntireColumn.AutoFit
    
End Function
Function Set_Font_Bold(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("D1") = "Font.Bold"
    For Each cell In Range("D2:D10")
        If cell.Row() < 6 Then
            cell.Value = "Font.Bold " & "True"
        Else
            cell.Value = "Font.Bold " & "False"
        End If
    Next cell
    ws.Range("D2:D10").Font.Bold = True
    ws.Range("D6:D10").Font.Bold = False
    ws.Range("D1:D10").EntireColumn.AutoFit
    
End Function
Function Set_Font_Italic(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("E1") = "Font.Italic"
    For Each cell In Range("E2:E10")
        If cell.Row() < 6 Then
            cell.Value = "Font.Italic " & "True"
        Else
            cell.Value = "Font.Italic " & "False"
        End If
    Next cell
    ws.Range("E2:E10").Font.Italic = True
    ws.Range("E6:E10").Font.Italic = False
    ws.Range("E1:E10").EntireColumn.AutoFit
    
End Function
Function Set_Font_Color(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("F1") = "Font.Color"
    For Each cell In Range("F2:F10")
        If cell.Row() < 4 Then
            cell.Value = "Font.Color " & "vbRed"
        ElseIf cell.Row() < 6 Then
            cell.Value = "Font.Color " & "vbBlue"
        ElseIf cell.Row() < 8 Then
            cell.Value = "Font.Color " & "vbGreen"
        Else
            cell.Value = "Font.Color " & "RGB(255, 0, 255)"
        End If
    Next cell
    ws.Range("F2:F10").Font.Color = vbRed
    ws.Range("F4:F10").Font.Color = vbBlue
    ws.Range("F6:F10").Font.Color = vbGreen
    ws.Range("F8:F10").Font.Color = RGB(255, 0, 255)
    ws.Range("F1:F10").EntireColumn.AutoFit
    
End Function
Function Set_Font_Underline(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("G1") = "Font.Underline"
    For Each cell In Range("G2:G10")
        If cell.Row() < 4 Then
            cell.Value = "Underline " & "xlSingle"
        ElseIf cell.Row() < 6 Then
            cell.Value = "Underline " & "xlUnderlineStyleNone"
        ElseIf cell.Row() < 8 Then
            cell.Value = "Underline " & "xlDouble"
        Else
            cell.Value = "Underline " & "xlUnderlineStyleNone"
        End If
    Next cell
    ws.Range("G2:G10").Font.Underline = xlSingle
    ws.Range("G4:G10").Font.Underline = xlUnderlineStyleNone
    ws.Range("G6:G10").Font.Underline = xlDouble
    ws.Range("G8:G10").Font.Underline = xlUnderlineStyleNone
    ws.Range("G1:G10").EntireColumn.AutoFit
    
End Function
Function Set_Font_Strikethrough(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("H1") = "Font.Strikethrough"
    For Each cell In Range("H2:H10")
        If cell.Row() < 6 Then
            cell.Value = "Strikethrough " & "True"
        Else
            cell.Value = "Strikethrough " & "False"
        End If
    Next cell
    ws.Range("H2:H10").Font.Strikethrough = True
    ws.Range("H6:H10").Font.Strikethrough = False
    ws.Range("H1:H10").EntireColumn.AutoFit
        
End Function
Function Set_Interior_Color(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("I1") = "Interior.Color"
    For Each cell In Range("I2:I10")
        If cell.Row() < 4 Then
            cell.Value = "Interior.Color " & "vbRed"
        ElseIf cell.Row() < 6 Then
            cell.Value = "Interior.Color " & "vbBlue"
        ElseIf cell.Row() < 8 Then
            cell.Value = "Interior.Color " & "vbGreen"
        Else
            cell.Value = "Interior.Color " & "RGB(255, 0, 255)"
        End If
    Next cell
    ws.Range("I2:I10").Interior.Color = vbRed
    ws.Range("I4:I10").Interior.Color = vbBlue
    ws.Range("I6:I10").Interior.Color = vbGreen
    ws.Range("I8:I10").Interior.Color = RGB(255, 0, 255)
    ws.Range("I1:I10").EntireColumn.AutoFit
    
End Function
Function Set_Phonetics(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("J1") = "Phonetics.Visible"
    For Each cell In Range("J2:J10")
        If cell.Row() < 4 Then
            cell.Value = "Phonetics.Alignment " & "xlPhoneticAlignLeft"
        ElseIf cell.Row() < 6 Then
            cell.Value = "Phonetics.Alignment " & "xlPhoneticAlignDistributed"
        Else
            cell.Value = "Phonetics.Alignment " & "xlPhoneticAlignCenter"
        End If
    Next cell
    ws.Range("J2:J10").Phonetics.Visible = True
    Range("J2:J10").Characters(1, 3).PhoneticCharacters = "Phonetics " & "X" & " Phonetics"
    Range("J2:J10").Phonetics.Alignment = xlPhoneticAlignLeft
    Range("J4:J10").Phonetics.Alignment = xlPhoneticAlignDistributed
    Range("J8:J10").Phonetics.Alignment = xlPhoneticAlignCenter
    ws.Range("J1:J10").EntireColumn.AutoFit
    
End Function

Function Set_ClearContents(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("K1") = "ClearContents"
    For Each cell In Range("K2:k10")
        If cell.Row() < 6 Then
            cell.Value = "ClearContents "
        Else
            cell.Value = "ClearContents "
        End If
    Next cell

    Range("K6:K10").ClearContents
    
    ws.Range("K1:K10").EntireColumn.AutoFit
    
End Function

