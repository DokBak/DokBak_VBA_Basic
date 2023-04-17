Attribute VB_Name = "Sheet_Class"
Option Explicit

Sub ThisWorkbookCreateSheets()

    Dim ws As Worksheet
    'pattern 1
    Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    ws.Name = "First Sheet"
    
    'pattern 2
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "Last Sheet"
    
    'pattern 3
    Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets("Last Sheet"))
    ws.Name = "Last Sheet Before"
    
    'pattern 4
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Last Sheet"))
    ws.Name = "Last Sheet After"
    
End Sub


