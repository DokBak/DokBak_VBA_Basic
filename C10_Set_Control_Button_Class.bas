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
    Dim Control_Button As Shape
    Dim Control_Button_Name As String
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

Function Set_Add_Labels(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    Worksheets(SheetName).Activate
    
    ActiveSheet.Labels.Add(50, 70, 150, 30).Select 'X_Position, Y_Position, X_Size, Y_Size
    'ActiveSheet.Shapes.Range(Array("Check Box 1")).Select 'Check Box Select
    Selection.Caption = "Label_Text"
    
End Function

Function Set_Add_CheckBoxes(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    Worksheets(SheetName).Activate
    
    ActiveSheet.CheckBoxes.Add(50, 120, 150, 30).Select 'X_Position, Y_Position, X_Size, Y_Size
    'ActiveSheet.Shapes.Range(Array("Check Box 1")).Select 'Check Box Select
    Selection.Caption = "CB1"
    Selection.Value = False
    ActiveSheet.CheckBoxes.Add(200, 120, 150, 30).Select 'X_Position, Y_Position, X_Size, Y_Size
    'ActiveSheet.Shapes.Range(Array("Check Box 2")).Select 'Check Box Select
    Selection.Caption = "CB2"
    Selection.Value = True
    
End Function
