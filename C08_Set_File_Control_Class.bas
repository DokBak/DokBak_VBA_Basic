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

Function Set_TextFile_NewLine_CRLF_Write()

    Dim FilePath As String
    Dim FileNumber As Integer
    Dim TextString As String
    
    FilePath = ThisWorkbook.Path & "/test.txt"
    
    FileNumber = FreeFile()
    Open FilePath For Append As #FileNumber
    
    TextString = "NewLine_CRLF" & vbCrLf & "Zero Line" & vbCrLf & "First Line" & vbCrLf & "Second Line" & vbCrLf
    Print #FileNumber, TextString
    
    'Linux Command -> od -c test.txt -> ¡Ír¡Ín
    
    Close #FileNumber
    
End Function

Function Set_TextFile_NewLine_LF_Write()

    Dim FilePath As String
    Dim FileNumber As Integer
    Dim TextString As String
    
    FilePath = ThisWorkbook.Path & "/test.txt"
    
    FileNumber = FreeFile()
    Open FilePath For Append As #FileNumber
    
    TextString = "NewLine_LF" & vbLf & "Zero Line" & vbLf & "First Line" & vbLf & "Second Line" & vbLf
    Print #FileNumber, TextString
    
    'Linux Command -> od -c test.txt -> ¡Ín
    
    Close #FileNumber
    
End Function

Function Set_TextFile_NewLine_CR_Write()

    Dim FilePath As String
    Dim FileNumber As Integer
    Dim TextString As String
    
    FilePath = ThisWorkbook.Path & "/test.txt"
    
    FileNumber = FreeFile()
    Open FilePath For Append As #FileNumber
    
    TextString = "NewLine_CR" & vbCr & "Zero Line" & vbCr & "First Line" & vbCr & "Second Line" & vbCr
    Print #FileNumber, TextString
    
    'Linux Command -> od -c test.txt -> ¡Ír
    
    Close #FileNumber
    
End Function
