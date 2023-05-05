Attribute VB_Name = "C01_Sheet_Class"
Option Explicit

Function ThisWorkbookCreateSheets(SheetName As String)

    Dim ws As Worksheet
    'pattern 1
    Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    ws.Name = SheetName & "First Sheet"
    
    'pattern 2
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = SheetName & "Last Sheet"
    
    ' Sample
    Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets("Main"))
    ws.Name = SheetName
    
    'pattern 3
    Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(SheetName))
    ws.Name = SheetName & " Sheet Before"
    
    'pattern 4
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(SheetName))
    ws.Name = SheetName & " Sheet After"
    
End Function

Function ThisWorkbookExists(SheetName As String)

    Dim ws As Worksheet
    Dim isSheetExists As Boolean
    
    isSheetExists = False
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = SheetName Then
            isSheetExists = True
            MsgBox "isSheetExists TRUE"
            Exit For
        End If
    Next ws
    
End Function

Function ChangeSheetColor(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Tab.Color = RGB(255, 0, 0)
    
    ws.Tab.ThemeColor = xlThemeColorDark1
    ws.Tab.TintAndShade = 0

End Function

Function HideSheets(SheetName As String)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Visible = xlSheetHidden

End Function

Function UnHideSheets(SheetName As String)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Visible = xlSheetVisible

End Function

Function ProtectSheet(SheetName As String)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)

    ws.Protect Password:="mypassword", DrawingObjects:=True, Contents:=True, Scenarios:=True
    'DrawingObjects : graphic objects, pictures, charts Permission
    'Contents : cell contents Permission
    
End Function

Function UnprotectSheet(SheetName As String)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Unprotect Password:="mypassword"
    
End Function

Function DeleteSheet(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    ws.Delete

End Function
