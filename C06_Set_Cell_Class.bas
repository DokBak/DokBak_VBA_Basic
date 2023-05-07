Attribute VB_Name = "C06_Set_Cell_AutoFilter_Class"
Option Explicit

Function Set_Cell_AutoFilter_Sheet(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    
    ws.Name = SheetName
    
    Worksheets(SheetName).Activate
    
    ws.Range("A1") = "No"
    For Each cell In Range("A2:A20")
        cell.Value = cell.Row() - 1
    Next cell
    
    ws.Range("B1") = "Random_Number"
    For Each cell In Range("B2:B20")
        cell.Value = "=int(Rand()*10) "
        cell.Copy
        cell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Next cell
    
    ws.Range("C1") = "Color"
    For Each cell In Range("C2:C20")
        If cell.Row() < 4 Then
            cell.Value = "Red"
        ElseIf cell.Row() < 10 Then
            cell.Value = "Green"
        ElseIf cell.Row() < 16 Then
            cell.Value = "Blue"
        Else
            cell.Value = "Black"
        End If
    Next cell
    
    ws.Range("C1") = "Fruit"
    For Each cell In Range("C2:C20")
        If cell.Row() < 2 Then
            cell.Value = "Apple"
        ElseIf cell.Row() < 5 Then
            cell.Value = "Orange"
        ElseIf cell.Row() < 8 Then
            cell.Value = "Melon"
        ElseIf cell.Row() < 11 Then
            cell.Value = "Peach"
        ElseIf cell.Row() < 13 Then
            cell.Value = "Blueberry"
        ElseIf cell.Row() < 17 Then
            cell.Value = "Strawberry"
        ElseIf cell.Row() < 19 Then
            cell.Value = "Grape"
        Else
            cell.Value = "Pear"
        End If
    Next cell
    
    
    ws.Range("D1") = "Test"
    For Each cell In Range("D2:D20")
        cell.Value = "Test"
    Next cell
    
End Function

Function Set_Cell_AutoFilter(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Worksheets(SheetName)
    ws.Activate
    
    ActiveSheet.AutoFilterMode = False
    ActiveSheet.Range("A1:D1").AutoFilter

End Function

Function Set_Cell_AutoFilter_Select(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    ws.Activate
    
    ActiveSheet.AutoFilterMode = False
    
    ActiveSheet.Range("A1:D1").AutoFilter Field:=3, Criteria1:="Strawberry" 'First Filter
    ActiveSheet.Range("A1:D1").AutoFilter Field:=1, Criteria1:="13"         'Second Filter
    
End Function

Function Set_Cell_AutoFilter_Clear(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    ws.Activate
    
    ActiveSheet.Range("A1:D1").AutoFilter Field:=3
    ActiveSheet.Range("A1:D1").AutoFilter Field:=1
    
End Function

Function Set_Cell_AutoFilter_xlAscending(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Worksheets(SheetName)
    ws.Activate
    
    ActiveSheet.AutoFilterMode = False
    ActiveSheet.Range("A1:D1").AutoFilter
    
    ws.AutoFilter.Sort.SortFields.Clear
    ws.AutoFilter.Sort.SortFields.Add2 _
        Key:=Range("B1"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ws.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Function

Function Set_Cell_AutoFilter_xlDescending(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Worksheets(SheetName)
    ws.Activate
    
    ActiveSheet.AutoFilterMode = False
    ActiveSheet.Range("A1:D1").AutoFilter
    
    ws.AutoFilter.Sort.SortFields.Clear
    ws.AutoFilter.Sort.SortFields.Add2 _
        Key:=Range("B1"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ws.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ws.AutoFilter.Sort.SortFields.Add2 _
        Key:=Range("C1"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ws.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Function

Function Set_Cell_UnAutoFilter(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets(SheetName)
    ws.Activate
    
    ActiveSheet.AutoFilterMode = False

End Function
