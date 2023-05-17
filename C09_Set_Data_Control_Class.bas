Attribute VB_Name = "C09_Set_Data_Control_Class"
Option Explicit

Function Set_Page_Layout_Sheet(SheetName As String)

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    
    ws.Name = SheetName
    
    Worksheets(SheetName).Activate
    
End Function

Function Set_RemoveDuplicates_No_Header(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    Worksheets(SheetName).Activate
    
    ws.Range("A1") = "AB"
    ws.Range("A2:A5") = "AA"
    ws.Range("A6") = "BB"
    ws.Range("A7:A8") = "aa"
    ws.Range("A9:A14") = "AB"
    ws.Range("A15:A19") = "AA"
    ws.Range("A20:A23") = "bA"
    ws.Range("A24") = "AB"
    ws.Range("B1") = "41"
    ws.Range("B2:B3") = "11"
    ws.Range("B4") = "23"
    ws.Range("B5:B6") = "41"
    ws.Range("B7:B9") = "42"
    ws.Range("B10:B13") = "47"
    ws.Range("B14:B16") = "11"
    ws.Range("B17:B18") = "23"
    ws.Range("B19:B20") = "42"
    ws.Range("B21:B24") = "45"
   
    ws.Range("$A$1:$B$24").RemoveDuplicates Columns:=1, Header:=xlNo
    
End Function

Function Set_RemoveDuplicates_Header(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    Worksheets(SheetName).Activate
    
    ws.Range("A1") = "AB"
    ws.Range("A2:A5") = "AA"
    ws.Range("A6") = "BB"
    ws.Range("A7:A8") = "aa"
    ws.Range("A9:A14") = "AB"
    ws.Range("A15:A19") = "AA"
    ws.Range("A20:A23") = "bA"
    ws.Range("A24") = "AB"
    ws.Range("B1") = "41"
    ws.Range("B2:B3") = "11"
    ws.Range("B4") = "23"
    ws.Range("B5:B6") = "41"
    ws.Range("B7:B9") = "42"
    ws.Range("B10:B13") = "47"
    ws.Range("B14:B16") = "11"
    ws.Range("B17:B18") = "23"
    ws.Range("B19:B20") = "42"
    ws.Range("B21:B24") = "45"
   
    ws.Range("$A$1:$B$24").RemoveDuplicates Columns:=2, Header:=xlYes
    
End Function

