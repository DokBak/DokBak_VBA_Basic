Attribute VB_Name = "M01_commonFunction"
Option Explicit

Function CreateSheetIfNotExists(SheetName As String) As Worksheet

    Dim ws As Worksheet
    Dim isSheetExists As Boolean
    Dim newSheet As Worksheet
    
    isSheetExists = False

    'Sheet Check
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = SheetName Then
            isSheetExists = True
            Exit For
        End If
    Next ws
    
    'Create New Sheet
    If Not isSheetExists Then
        If SheetName = "VersionHistory" Then
            Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(1))
        Else
            Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        End If
        newSheet.Name = SheetName
        Set CreateSheetIfNotExists = newSheet
    Else
        Set CreateSheetIfNotExists = Nothing
    End If
    
End Function
Function ResizeColumnsInSheet(SheetName As String)

    Dim ws As Worksheet
    Dim lastColumn As Long
    Dim i As Long
     
    Worksheets(SheetName).Activate
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        ResizeColumnsInSheet = False
        Exit Function
    End If
    
    Columns("A:XFD").Select
    Selection.ColumnWidth = 3
    Selection.Interior.Color = RGB(255, 255, 255) 'white
    Selection.NumberFormat = "@" 'General:General, @:string , 0.00 :Double
End Function

Function VersionHistorySet(SheetName As String)
    
    Dim ws As Worksheet
    Dim rng As Range
    
    Worksheets(SheetName).Activate
    
    Set ws = ThisWorkbook.Sheets(SheetName)
    'ич
    'Cell Merge
    ws.Range("B2:V3").Merge 'FileName Column
    ws.Range("W2:AA3").Merge 'Version Column
    ws.Range("AB2:AF3").Merge 'ModifiedDate Column
    ws.Range("AG2:AK3").Merge 'CreateUser Column
    ws.Range("AL2:AP3").Merge 'ModifiedUser Column
    
    'Cell Value : Center
    ws.Range("B2:V3").HorizontalAlignment = xlCenter
    ws.Range("W2:AA3").HorizontalAlignment = xlCenter
    ws.Range("AB2:AF3").HorizontalAlignment = xlCenter
    ws.Range("AG2:AK3").HorizontalAlignment = xlCenter
    ws.Range("AL2:AP3").HorizontalAlignment = xlCenter
    
    'Cell Color
    ws.Range("B2").Interior.Color = RGB(102, 315, 102)
    ws.Range("W2").Interior.Color = RGB(102, 315, 102)
    ws.Range("AB2").Interior.Color = RGB(102, 315, 102)
    ws.Range("AG2").Interior.Color = RGB(102, 315, 102)
    ws.Range("AL2").Interior.Color = RGB(102, 315, 102)

   'Cell Border
    With ws.Range("B2:AP3").Borders
        .LineStyle = xlContinuous 'line
        .Color = RGB(0, 0, 0) ' black
        .Weight = xlThin 'xlThin, xlMedium
    End With
    
    'Cell Value
    ws.Range("B2").Value = "FileName"
    ws.Range("W2").Value = "Version"
    ws.Range("AB2").Value = "ModifiedDate"
    ws.Range("AG2").Value = "CreateUser"
    ws.Range("AL2").Value = "ModifiedUser"
    
    'иш
    'Cell Merge
    ws.Range("B4:V4").Merge 'FileName Column
    ws.Range("W4:AA4").Merge 'Version Column
    ws.Range("AB4:AF4").Merge 'ModifiedDate Column
    ws.Range("AG4:AK4").Merge 'CreateUser Column
    ws.Range("AL4:AP4").Merge 'ModifiedUser Column
    
    'Cell Value : Center
    ws.Range("B4").HorizontalAlignment = xlCenter
    ws.Range("W4").HorizontalAlignment = xlCenter
    ws.Range("AB4").HorizontalAlignment = xlCenter
    ws.Range("AG4").HorizontalAlignment = xlCenter
    ws.Range("AL4").HorizontalAlignment = xlCenter
    
    'Cell Color
    ws.Range("B4").Interior.Color = RGB(128, 128, 128)
    ws.Range("W4").Interior.Color = RGB(128, 128, 128)
    ws.Range("AB4").Interior.Color = RGB(128, 128, 128)
    ws.Range("AG4").Interior.Color = RGB(128, 128, 128)
    ws.Range("AL4").Interior.Color = RGB(128, 128, 128)
    
    'Cell Border
    With ws.Range("B4:AP4").Borders
        .LineStyle = xlContinuous 'line
        .Color = RGB(0, 0, 0) ' black
        .Weight = xlThin 'xlThin, xlMedium
    End With
    
    'Cell Format
    ws.Range("B4").NumberFormat = "General"
    ws.Range("W4").NumberFormat = "0.0"
    ws.Range("AB4").NumberFormat = "YYYY/MM/DD"
    ws.Range("AG4").NumberFormat = "General"
    ws.Range("AL4").NumberFormat = "General"
    
    'Cell Value
    ws.Range("B4").Value = ThisWorkbook.Name
    ws.Range("W4").Formula = "=LOOKUP(1,0/(D:D<>""""),D:D)"
    ws.Range("AB4").Formula = "=LOOKUP(1,0/(F:F<>""""),F:F)"
    ws.Range("AG4").Formula = "=IF(AL8="""","""",AL8)"
    ws.Range("AL4").Formula = "=LOOKUP(1,0/(AL:AL<>""""),AL:AL)"
    
    'ищ
    'Cell Merge
    ws.Range("B6:C7").Merge 'No Column
    ws.Range("D6:E7").Merge 'Version Column
    ws.Range("F6:I7").Merge 'ModifiedDate Column
    ws.Range("J6:M7").Merge 'ModifiedReason Column
    ws.Range("N6:S7").Merge 'ModifiedArea Column
    ws.Range("T6:AK7").Merge 'ModifiedContents Column
    ws.Range("AL6:AP7").Merge 'ModifiedUser Column
    
    'Cell Value : Center
    ws.Range("B6").HorizontalAlignment = xlCenter
    ws.Range("D6").HorizontalAlignment = xlCenter
    ws.Range("F6").HorizontalAlignment = xlCenter
    ws.Range("J6").HorizontalAlignment = xlCenter
    ws.Range("N6").HorizontalAlignment = xlCenter
    ws.Range("T6").HorizontalAlignment = xlCenter
    ws.Range("AL6").HorizontalAlignment = xlCenter
    
    'Cell Color
    ws.Range("B6").Interior.Color = RGB(102, 315, 102)
    ws.Range("D6").Interior.Color = RGB(102, 315, 102)
    ws.Range("F6").Interior.Color = RGB(102, 315, 102)
    ws.Range("J6").Interior.Color = RGB(102, 315, 102)
    ws.Range("N6").Interior.Color = RGB(102, 315, 102)
    ws.Range("T6").Interior.Color = RGB(102, 315, 102)
    ws.Range("AL6").Interior.Color = RGB(102, 315, 102)
    
    'Cell Border
    With ws.Range("B6:AP7").Borders
        .LineStyle = xlContinuous 'line
        .Color = RGB(0, 0, 0) ' black
        .Weight = xlThin 'xlThin, xlMedium
    End With
    
    'Cell Value
    ws.Range("B6").Value = "No."
    ws.Range("D6").Value = "Version"
    ws.Range("F6").Value = "ModifiedDate"
    ws.Range("J6").Value = "ModifiedReason"
    ws.Range("N6").Value = "ModifiedArea"
    ws.Range("T6").Value = "ModifiedContents"
    ws.Range("AL6").Value = "ModifiedUser"
    
    'иъ
    'Cell Merge
    ws.Range("B8:C8").Merge 'No Column
    ws.Range("D8:E8").Merge 'Version Column
    ws.Range("F8:I8").Merge 'ModifiedDate Column
    ws.Range("J8:M8").Merge 'ModifiedReason Column
    ws.Range("N8:S8").Merge 'ModifiedArea Column
    ws.Range("T8:AK8").Merge 'ModifiedContents Column
    ws.Range("AL8:AP8").Merge 'ModifiedUser Column
    
    'Cell Value : Center
    ws.Range("B8").HorizontalAlignment = xlCenter
    ws.Range("D8").HorizontalAlignment = xlCenter
    ws.Range("F8").HorizontalAlignment = xlCenter
    ws.Range("J8").HorizontalAlignment = xlCenter
    ws.Range("N8").HorizontalAlignment = xlCenter
    ws.Range("T8").HorizontalAlignment = xlCenter
    ws.Range("AL8").HorizontalAlignment = xlCenter
    
    'Cell Color
    ws.Range("B8").Interior.Color = RGB(128, 128, 128)
    ws.Range("D8").Interior.Color = RGB(128, 128, 128)
    
    'Cell Border
    With ws.Range("B8:AP8").Borders
        .LineStyle = xlContinuous 'line
        .Color = RGB(0, 0, 0) ' black
        .Weight = xlThin 'xlThin, xlMedium
    End With
    
    'Cell Format
    ws.Range("B8").NumberFormat = "0"
    ws.Range("D8").NumberFormat = "0.0"
    ws.Range("F8").NumberFormat = "YYYY/MM/DD"
    
    'Cell Value
    ws.Range("B8").Formula = "=row()-7"
    ws.Range("D8").Formula = "=if(B8<>"""",if(B8=1,1,D7+0.1),)"
    ws.Range("F8").Value = "2023/01/01"
    Range("J8:M8").Select
    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Validation!$B$3:$B$9"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    ws.Range("N8").Value = "Full"
    ws.Range("T8").Value = "Full Create"
    ws.Range("AL8").Value = "Jeong"
    
End Function

Function ValidationSet(SheetName As String)
    
    Dim ws As Worksheet
    Dim rng As Range
    
    Worksheets(SheetName).Activate
    
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    'ич
    'Cell Color
    ws.Range("B2").Interior.Color = RGB(102, 315, 102)
    
    'Cell Border
    With ws.Range("B2:B9").Borders
        .LineStyle = xlContinuous 'line
        .Color = RGB(0, 0, 0) ' black
        .Weight = xlThin 'xlThin, xlMedium
    End With
    
    'Cell Value
    ws.Range("B2").Value = "Modify_Reason"
    ws.Range("B3").Value = "New"
    ws.Range("B4").Value = "Macro_Create"
    ws.Range("B5").Value = "Macro_Modify"
    ws.Range("B6").Value = "Macro_Delete"
    ws.Range("B7").Value = "Sheet_Create"
    ws.Range("B8").Value = "Sheet_Modify"
    ws.Range("B9").Value = "Sheet_Delete"
    
    ws.Range("B:B").EntireColumn.AutoFit
    
End Function

Function SheetList(SheetName As String)
    Dim ws As Worksheet
    Dim wsList As String
    Dim wsName As String
    Dim i As Integer
    
    Worksheets(SheetName).Activate
    
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ws.Range("C:C").ClearContents
        
    'Cell Value
    ws.Range("C2").Value = "Sheet_List"
    'Cell Color
    ws.Range("C2").Interior.Color = RGB(102, 315, 102)
    
    With ws.Cells(2, 3).Borders
    .LineStyle = xlContinuous 'line
    .Color = RGB(0, 0, 0) ' black
    .Weight = xlThin 'xlThin, xlMedium
    End With
    
    i = 1
    While i <= ThisWorkbook.Sheets.Count
       Cells(i + 2, 3).Value = ThisWorkbook.Sheets(i).Name
       
           'Cell Border
         With ws.Cells(i + 2, 3).Borders
        .LineStyle = xlContinuous 'line
        .Color = RGB(0, 0, 0) ' black
        .Weight = xlThin 'xlThin, xlMedium
        End With
       
       i = i + 1
    Wend

    ws.Range("C:C").EntireColumn.AutoFit
    
End Function

Function SheetsInitialize()
    Dim wbkWorkBook As Workbook
    Set wbkWorkBook = ThisWorkbook
    Dim ws As Worksheet
    
    Dim i, Delete_Sheet_Number, All_Sheets_Count, Del_cnt As Integer
    All_Sheets_Count = wbkWorkBook.Worksheets.Count
    
    Dim Not_Delete_Sheet_Name As String
    Dim Del_Sheet As String
    
    Worksheets("Validation").Activate
    
    Not_Delete_Sheet_Name = "[Validation, VersionHistory, Main]"
    
    Del_cnt = 0
    
    'Dim rc As VbMsgBoxResult
    'rc = MsgBox("Sheets Initialize", vbYesNo + vbQuestion)
    
    'If rc = vbYes Then
        
        Dim newSheet As Worksheet
        Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        newSheet.Name = "DeleteSheet"
        Set ws = wbkWorkBook.Sheets("DeleteSheet")
        With ws
            For Delete_Sheet_Number = 1 To All_Sheets_Count
                If InStr(1, Not_Delete_Sheet_Name, wbkWorkBook.Sheets(Delete_Sheet_Number).Name) = 0 Then
                    Del_cnt = Del_cnt + 1
                    Worksheets("DeleteSheet").Cells(1, Del_cnt) = wbkWorkBook.Sheets(Delete_Sheet_Number).Name
                End If
            Next Delete_Sheet_Number
            
            For i = 1 To Del_cnt
                Application.DisplayAlerts = False
                Del_Sheet = Worksheets("DeleteSheet").Cells(1, i).Value
                
                wbkWorkBook.Sheets(Del_Sheet).Delete
            Next i
                  
        End With
    'End If
    
    wbkWorkBook.Sheets("DeleteSheet").Delete
End Function

Function HideSheet(SheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    Worksheets("Main").Activate
    
    ws.Visible = xlSheetHidden

End Function

Function ThisWorkbookExists(SheetName As String)

    Dim ws As Worksheet
    Dim isSheetExists As Boolean
    
    isSheetExists = False
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = SheetName Then
            isSheetExists = True
            DeleteSheet (SheetName)
            Exit For
        End If
    Next ws
    
End Function

Function DeleteSheet(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    ws.Delete

End Function

Sub MainInit()
    Dim SheetName As String
    
    'Validation Sheet Create
    SheetName = "Validation"
    ThisWorkbookExists (SheetName)
    CreateSheetIfNotExists (SheetName)
    ResizeColumnsInSheet (SheetName)
    ValidationSet (SheetName)
    
    'VersionHistory Sheet Create
    SheetName = "VersionHistory"
    ThisWorkbookExists (SheetName)
    CreateSheetIfNotExists (SheetName)
    ResizeColumnsInSheet (SheetName)
    VersionHistorySet (SheetName)
    
    'Main Sheet Create
    SheetName = "Main"
    ThisWorkbookExists (SheetName)
    CreateSheetIfNotExists (SheetName)
    ResizeColumnsInSheet (SheetName)
    
    'Default Sheets
    SheetsInitialize
    
    'Validation Sheet Add
    SheetName = "Validation"
    SheetList (SheetName)
    
    'SheetHide
    SheetName = "Validation"
    HideSheet (SheetName)
    
End Sub



