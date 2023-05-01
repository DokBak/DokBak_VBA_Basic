Attribute VB_Name = "T01_Test_Module"
Option Explicit


Sub TS01_Sheet_Class_Test()

    Dim SheetName As String
    SheetName = "Sheet_Class"
    Worksheets("Main").Activate
    
    ThisWorkbookCreateSheets (SheetName)
    ThisWorkbookExists (SheetName)
    ChangeSheetColor (SheetName)
    HideSheets (SheetName)
    UnHideSheets (SheetName)
    ProtectSheet (SheetName)
    UnprotectSheet (SheetName)
    
    Application.DisplayAlerts = False
    SheetName = "Sheet_Class"
    DeleteSheet (SheetName)
    SheetName = "Sheet_ClassFirst Sheet"
    DeleteSheet (SheetName)
    SheetName = "Sheet_Class Sheet Before"
    DeleteSheet (SheetName)
    SheetName = "Sheet_Class Sheet After"
    DeleteSheet (SheetName)
    SheetName = "Sheet_ClassLast Sheet"
    DeleteSheet (SheetName)
    
End Sub

Sub TS02_Set_Cell_Class_Test()

    Dim SheetName As String
    
    SheetName = "Set_Cell_Class"
    Set_Cell_Sheet (SheetName)
    Set_Font_Name (SheetName)
    Set_Font_Size (SheetName)
    Set_Font_Bold (SheetName)
    Set_Font_Italic (SheetName)
    Set_Font_Color (SheetName)
    Set_Font_Underline (SheetName)
    Set_Font_Strikethrough (SheetName)
    Set_Interior_Color (SheetName)
    Set_Phonetics (SheetName)
    Set_ClearContents (SheetName)
    Set_Borders (SheetName)
    DeleteSheet (SheetName)
    
End Sub
Sub TS03_Set_Align_Class_Test()

    Dim SheetName As String
    
    SheetName = "Set_Align_Class"
    Set_Align_Sheet (SheetName)
    Set_Horizontal_Alignment (SheetName)
    Set_Vertical_Alignment (SheetName)
    Set_IndentLevel (SheetName)
    Set_Orientation (SheetName)
    Set_WrapText (SheetName)
    Set_Merge (SheetName)
    Set_AutoFit (SheetName)
    DeleteSheet (SheetName)
    
End Sub
Sub TS04_Set_Fomat_Class_Test()
    
    Dim SheetName As String
    
    SheetName = "Set_Format_Class"
    Set_Format_Sheet (SheetName)
    Set_NumberFormat_Format (SheetName)
    Set_Data_Format (SheetName)
    Set_Example_Format (SheetName)
    DeleteSheet (SheetName)
    
End Sub
Sub TS05_Set_Conditional_Class_Test()
    
    Dim SheetName As String
    
    SheetName = "Set_Conditional_Class"
    Set_Conditional_Sheet (SheetName)
    Set_xlCellValue (SheetName)
    Set_xlTextString (SheetName)
    Set_xlTimePeriod (SheetName)
    Set_xlBlanksCondition (SheetName)
    DeleteSheet (SheetName)
    
End Sub
