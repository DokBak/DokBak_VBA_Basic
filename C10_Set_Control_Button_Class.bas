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

Function Set_Add_OptionButtons(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    Worksheets(SheetName).Activate
    
    ActiveSheet.OptionButtons.Add(50, 170, 150, 30).Select 'X_Position, Y_Position, X_Size, Y_Size
    'ActiveSheet.Shapes.Range(Array("Option Button 1")).Select 'Option Button Select
    Selection.Caption = "OB1"
    Selection.Value = True
    ActiveSheet.OptionButtons.Add(200, 170, 150, 30).Select 'X_Position, Y_Position, X_Size, Y_Size
    'ActiveSheet.Shapes.Range(Array("Option Button 2")).Select 'Option Button Select
    Selection.Caption = "OB2"
    Selection.Value = True
    
End Function

Function Set_Add_ListBoxes(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    Worksheets(SheetName).Activate
    
    ActiveSheet.ListBoxes.Add(50, 220, 150, 30).Select 'X_Position, Y_Position, X_Size, Y_Size
    'ActiveSheet.Shapes.Range(Array("List Box 1")).Select 'ListBox Select
    Selection.AddItem "Apple"
    Selection.AddItem "Banana"
    Selection.AddItem "Peach"
    Selection.AddItem "Orange"
    Selection.AddItem "Melon"
    ActiveSheet.ListBoxes.Add(250, 220, 150, 30).Select 'X_Position, Y_Position, X_Size, Y_Size
    'ActiveSheet.Shapes.Range(Array("List Box 2")).Select 'ListBox Select
    'Selection.ColumnCount = 3
    Selection.AddItem "List1"
    Selection.AddItem "List2"
    Selection.AddItem "List3"
    Selection.AddItem "List4"
    Selection.AddItem "List5"
    ActiveSheet.ListBoxes.Add(450, 220, 150, 30).Select 'X_Position, Y_Position, X_Size, Y_Size
    'ActiveSheet.Shapes.Range(Array("List Box 3")).Select 'ListBox Select
    Selection.MultiSelect = fmMultiSelectMulti
    Selection.AddItem "Box1"
    Selection.AddItem "Box2"
    Selection.AddItem "Bxo3"
    Selection.AddItem "Box4"
    Selection.AddItem "Box5"
    
End Function

Function Set_Add_DropDowns(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    Worksheets(SheetName).Activate
    
    ActiveSheet.DropDowns.Add(50, 270, 150, 30).Select 'X_Position, Y_Position, X_Size, Y_Size
    'ActiveSheet.Shapes.Range(Array("DropDowns 1")).Select 'DropDowns Select
    Selection.AddItem "Box1"
    Selection.AddItem "Box2"
    Selection.AddItem "Box3"
    Selection.AddItem "Box4"
    Selection.AddItem "Box5"
    Selection.AddItem "Box6"
    Selection.AddItem "Box7"
    Selection.AddItem "Box8"
    Selection.AddItem "Box9"
    Selection.AddItem "Box10"
    Selection.AddItem "Box11"
    
End Function

Function Set_Add_ScrollBars(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    Worksheets(SheetName).Activate
    
    ActiveSheet.ScrollBars.Add(50, 320, 150, 30).Select 'X_Position, Y_Position, X_Size, Y_Size
    'ActiveSheet.Shapes.Range(Array("ScrollBars 1")).Select 'ScrollBars Select

End Function
