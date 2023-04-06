Attribute VB_Name = "Module1"
Option Explicit

Sub ExcelVariables()
    Dim VariableByte As Byte
    Dim VariableInt As Integer
    Dim VariableLong As Long
    Dim VariableSingle As Single
    Dim VariableDouble As Double
    Dim VariableBoolean As Boolean
    Dim VariableString As String
    Dim VariableDate As Date
    Dim VariableObject As Object
    
    VariableByte = 1
    VariableInt = 2
    VariableLong = 3
    VariableSingle = 4
    VariableDouble = 5
    VariableBoolean = True
    VariableString = "TestString"
    VariableDate = "2023/02/20"
    
    MsgBox VariableByte & "," & VariableInt & "," & VariableLong & "," & VariableSingle & vbTab & VariableDouble & vbCr & VariableBoolean & vbCr & VariableString & vbCr & VariableDate
    MsgBox "Test" & vbBack
    MsgBox vbNullChar
    
    Const ConstVariableInt = 11
    
    MsgBox ConstVariableInt
End Sub
Sub ArrayVariables()
    Dim VariableArray(1 To 4) As Integer
    Dim VariableArray2(3) As Integer
    VariableArray(1) = 1
    VariableArray(2) = 2
    VariableArray(3) = 3
    VariableArray(4) = 4
    VariableArray2(0) = 5
    VariableArray2(1) = 6
    VariableArray2(2) = 7
    VariableArray2(3) = 8

    MsgBox VariableArray(1) & "," & VariableArray(2) & "," & VariableArray(3) & "," & VariableArray(4) & "," & VariableArray2(0) & "," & VariableArray2(1) & "," & VariableArray2(2) & "," & VariableArray2(3)
End Sub

Sub Current_File_Path()
    Dim wbkDir As String
    Dim wbkName As String
    
    wbkDir = ThisWorkbook.Path
    wbkName = ThisWorkbook.Name
    MsgBox wbkDir & "\" & wbkName
End Sub
