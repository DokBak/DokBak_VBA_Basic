Attribute VB_Name = "C08_Set_File_Control_Class"
Option Explicit

Function Set_File_Control_Sheet(SheetName As String)

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    
    ws.Name = SheetName
    
    Worksheets(SheetName).Activate
    
End Function

Function Set_Basic_TextFile_Write()

    Dim FilePath As String
    Dim FileNumber As Integer
    
    FilePath = ThisWorkbook.Path & "/test.txt"
    
    FileNumber = FreeFile()
    Open FilePath For Output As #FileNumber
    
    Print #FileNumber, "Hello, world!"
    Print #FileNumber, "This is a test."
    
    Close #FileNumber
    
End Function

Function Set_Ouptut_TextFile_OverWrite()

    Dim FilePath As String
    Dim FileNumber As Integer
    
    FilePath = ThisWorkbook.Path & "/test.txt"
    
    FileNumber = FreeFile()
    Open FilePath For Output As #FileNumber
    
    Print #FileNumber, "OverWrite Mode"
    Print #FileNumber, "For Output"
    Print #FileNumber, "First Line"
    Print #FileNumber, "Second Line"
    
    Close #FileNumber
    
End Function

Function Set_Append_TextFile_OverWrite()

    Dim FilePath As String
    Dim FileNumber As Integer
    
    FilePath = ThisWorkbook.Path & "/test.txt"
    
    FileNumber = FreeFile()
    Open FilePath For Append As #FileNumber
    
    Print #FileNumber, "Append_Write Mode"
    Print #FileNumber, "For Append"
    Print #FileNumber, "First Line"
    Print #FileNumber, "Second Line"
    
    Close #FileNumber
    
End Function

