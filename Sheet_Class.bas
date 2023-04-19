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

Sub ChangeSheetColor()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Main")
    
    ws.Tab.Color = RGB(255, 0, 0)
    
    ws.Tab.ThemeColor = xlThemeColorDark1
    ws.Tab.TintAndShade = 0

End Sub

Sub HideSheet()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Main")
    
    ws.Visible = xlSheetHidden

End Sub

Sub UnHideSheet()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Main")
    
    ws.Visible = xlSheetVisible

End Sub

Sub ProtectSheet()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Main")

    ws.Protect Password:="mypassword", DrawingObjects:=True, Contents:=True, Scenarios:=True
    'DrawingObjects :도형, 그림, 차트 등의 그래픽 객체를 수정 권한설정
    'Contents :셀의 내용의 수정 권한설정
    
End Sub

Sub UnprotectSheet()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Main")
    
    ws.Unprotect Password:="mypassword"
    
End Sub

