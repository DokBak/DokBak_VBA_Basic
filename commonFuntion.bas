Attribute VB_Name = "commonFunction"
Option Explicit

Function CreateSheetIfNotExists(sheetName As String) As Worksheet

    Dim ws As Worksheet
    Dim isSheetExists As Boolean
    Dim newSheet As Worksheet
    
    isSheetExists = False
    
    'Sheet Check
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = sheetName Then
            isSheetExists = True
            Exit For
        End If
    Next ws
    
    'Create New Sheet
    If Not isSheetExists Then
        Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        newSheet.Name = sheetName
        Set CreateSheetIfNotExists = newSheet
    Else
        Set CreateSheetIfNotExists = Nothing
    End If
    
End Function
Function ResizeColumnsInSheet(sheetName As String)

    Dim ws As Worksheet
    Dim lastColumn As Long
    Dim i As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
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

Function VersionHistorySet(sheetName As String)
    
    Dim ws As Worksheet
    Dim rng As Range
    
    Set ws = ThisWorkbook.Sheets(sheetName)
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

Function ValidationSet(sheetName As String)
    
    Dim ws As Worksheet
    Dim rng As Range
    
    Set ws = ThisWorkbook.Sheets(sheetName)
    
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
    
End Function
Function sheetList(sheetName As String)
    Dim ws As Worksheet
    Dim wsList As String
    Dim wsName As String
    Dim i As Integer
     
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    'Cell Value
    ws.Range("C2").Value = "Sheet_List"
    
    i = 1
    While i <= ThisWorkbook.Sheets.Count

        'For Each wbs In ThisWorkbook.Sheets
           Cells(i + 2, 3).Value = ThisWorkbook.Sheets(i).Name
           i = i + 1
        'Next wbs
        
    Wend

End Function

Function DeleteSheet(sheetName As String)
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
End Function

Function SheetListReset()
    
    Dim wbkSheet As Integer
    Dim myList As Variant
    Dim myValue As Variant
    Dim i As Integer
    Dim found As Boolean

    myList = Array("VersionHistory", "Validation")

    For i = LBound(myList) To UBound(myList)
        wbkSheet = 1
        While wbkSheet <= ThisWorkbook.Sheets.Count
            If myList(i) <> ThisWorkbook.Sheets(wbkSheet).Name Then
            
                DeleteSheet (ThisWorkbook.Sheets(wbkSheet).Name)
            Else
                wbkSheet = wbkSheet + 1
            End If
        Wend
        
    Next i

End Function



Sub VersionHistorySheet()
    Dim sheetName As String
    
    sheetName = "Validation"
    CreateSheetIfNotExists (sheetName)
    ResizeColumnsInSheet (sheetName)
    ValidationSet (sheetName)
    sheetList (sheetName)
    
    sheetName = "VersionHistory"
    CreateSheetIfNotExists (sheetName)
    ResizeColumnsInSheet (sheetName)
    VersionHistorySet (sheetName)
        
End Sub

