Attribute VB_Name = "commonFunction"
Option Explicit

Function CreateSheetIfNotExists(sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim isSheetExists As Boolean
    Dim newSheet As Worksheet
    
    isSheetExists = False
    
    'Sheet Check
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = sheetName Then
            isSheetExists = True
            Exit For
        End If
    Next ws
    
    'Create New Sheet
    If Not isSheetExists Then
        Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        newSheet.Name = sheetName
        MsgBox "Create New Sheet"
        Set CreateSheetIfNotExists = newSheet
    Else
        MsgBox "Exist Sheet"
        Set CreateSheetIfNotExists = Nothing
    End If
End Function

