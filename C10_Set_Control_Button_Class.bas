Attribute VB_Name = "C10_Set_Control_Button_Class"
Option Explicit

Function Set_Control_Button_Sheet(SheetName As String)

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    
    ws.Name = SheetName
    
    Worksheets(SheetName).Activate
    
End Function

Sub MsgTest()
    MsgBox "Test VBA"
End Sub

Function Set_Add_Button(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    Worksheets(SheetName).Activate
    
    ActiveSheet.Buttons.Add(50, 20, 150, 30).Select 'X_Position, Y_Position, X_Size, Y_Size
    ActiveSheet.Shapes("Button 1").Name = "Control_Button_Name" 'Button Name Change
    ActiveSheet.Shapes.Range(Array("Control_Button_Name")).Select 'Button Select
    Selection.Characters.text = "Control_Button_Text" 'Button Text Change
    Selection.OnAction = "MsgTest" 'VBA Function or Sub
    With Selection.Characters(Start:=1, Length:=3).Font
        .Name = "Arial"
        .FontStyle = "Normal"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    
End Function


