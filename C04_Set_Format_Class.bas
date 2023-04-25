Attribute VB_Name = "C04_Set_Format_Class"
Option Explicit

Function Set_Format_Sheet(SheetName As String)

    Dim ws As Worksheet
    Dim cell As Variant
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    
    ws.Name = SheetName
    
    Worksheets(SheetName).Activate
    
    ws.Range("A1") = "Format"
    For Each cell In Range("A2:A10")
        If cell.Row() < 3 Then
            cell.Value = "Integer Type"
        ElseIf cell.Row() < 4 Then
            cell.Value = "Float Type"
        ElseIf cell.Row() < 5 Then
            cell.Value = "Fractional Type"
        ElseIf cell.Row() < 6 Then
            cell.Value = "Date Type : YYYYMMDD"
        ElseIf cell.Row() < 7 Then
            cell.Value = "Time Type : hh:mm:ss"
        ElseIf cell.Row() < 8 Then
            cell.Value = "Currency Type"
        ElseIf cell.Row() < 9 Then
            cell.Value = "Accounting Tpye"
        ElseIf cell.Row() < 10 Then
            cell.Value = "String Type"
        Else
            cell.Value = "Custom Type"
        End If
    Next cell
    
    ws.Range("A1:A10").EntireColumn.AutoFit
    
End Function


