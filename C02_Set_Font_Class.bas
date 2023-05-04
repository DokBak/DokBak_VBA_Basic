Attribute VB_Name = "C02_Set_Font_Class"
Option Explicit

Function Set_Font_Sheet(SheetName As String)
    
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
    ws.Range("J2:J10").Characters(1, 3).PhoneticCharacters = "Phonetics " & "X" & " Phonetics"
    ws.Range("J2:J10").Phonetics.Alignment = xlPhoneticAlignLeft
    ws.Range("J4:J10").Phonetics.Alignment = xlPhoneticAlignDistributed
    ws.Range("J8:J10").Phonetics.Alignment = xlPhoneticAlignCenter
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

    ws.Range("K6:K10").ClearContents
    
    ws.Range("K1:K10").EntireColumn.AutoFit
    
End Function

Function Set_Borders(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("L1") = "Borders"
    For Each cell In Range("L2:L27")
        If cell.Row() < 3 Then
            cell.Value = "Borders(xlEdgeTop).LineStyle = xlContinuous "
        ElseIf cell.Row() < 4 Then
            cell.Value = "Borders(xlEdgeBottom).LineStyle = xlDash "
        ElseIf cell.Row() < 5 Then
            cell.Value = "Borders(xlEdgeLeft).LineStyle = xlDashDot "
        ElseIf cell.Row() < 6 Then
            cell.Value = "Borders(xlEdgeRight).LineStyle = xlDashDotDot "
        ElseIf cell.Row() < 8 Then
            cell.Value = "Borders(xlInsideVertical).LineStyle = xlDot "
        ElseIf cell.Row() < 9 Then
            cell.Value = "Borders(xlInsideHorizontal).LineStyle = xlDouble "
        ElseIf cell.Row() < 10 Then
            cell.Value = "Borders(xlDiagonalDown).LineStyle = xlContinuous "
        ElseIf cell.Row() < 11 Then
            cell.Value = "Borders(xlDiagonalDown).LineStyle = xlLineStyleNone "
        ElseIf cell.Row() < 13 Then
            cell.Value = "Borders.Weight = xlThick "
        ElseIf cell.Row() = 13 Then
        ElseIf cell.Row() < 16 Then
            cell.Value = "Borders.Weight = xlMedium "
        ElseIf cell.Row() = 16 Then
        ElseIf cell.Row() < 19 Then
            cell.Value = "Borders.Weight = xlThick "
        ElseIf cell.Row() = 19 Then
        ElseIf cell.Row() < 22 Then
            cell.Value = "Borders.Weight = xlHairline "
        ElseIf cell.Row() = 22 Then
        ElseIf cell.Row() < 25 Then
            cell.Value = "Borders.Color = RGB(255, 0, 0) "
        ElseIf cell.Row() = 25 Then
        Else
            cell.Value = "Borders.Color = xlRed"
        End If
    Next cell

    ws.Range("L2:M2").Borders(xlEdgeTop).LineStyle = xlContinuous
    ws.Range("L3:M3").Borders(xlEdgeBottom).LineStyle = xlDash
    ws.Range("L4:M4").Borders(xlEdgeLeft).LineStyle = xlDashDot
    ws.Range("L5:M5").Borders(xlEdgeRight).LineStyle = xlDashDotDot
    ws.Range("L6:M6").Borders(xlInsideVertical).LineStyle = xlDot
    ws.Range("L7:M8").Borders(xlInsideHorizontal).LineStyle = xlDouble
    ws.Range("L9:M9").Borders(xlDiagonalDown).LineStyle = xlContinuous
    ws.Range("L10:M10").Borders(xlDiagonalDown).LineStyle = xlLineStyleNone
    
    ws.Range("L11:M12").Borders.Weight = xlThick
    ws.Range("L14:M15").Borders.Weight = xlMedium
    ws.Range("L17:M18").Borders.Weight = xlThick
    ws.Range("L20:M21").Borders.Weight = xlHairline
    
    ws.Range("L23:M24").Borders.Weight = xlThick
    ws.Range("L23:M24").Borders.Color = RGB(255, 0, 255)
    
    ws.Range("L26:M27").Borders.Weight = xlThick
    ws.Range("L26:M27").Borders.Color = vbRed
    ws.Range("L1:M27").EntireColumn.AutoFit
    
End Function


