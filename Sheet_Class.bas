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
Sub ThisWorkbookExists()

    Dim ws As Worksheet
    Dim isSheetExists As Boolean
    Dim checkSheet As String
    
    checkSheet = "checkSheetName"
    isSheetExists = False
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = checkSheet Then
            isSheetExists = True
            Exit For
        End If
    Next ws
    
End Sub

