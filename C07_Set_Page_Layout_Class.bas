Attribute VB_Name = "C07_Set_Page_Layout_Class"
Option Explicit

Function Set_Page_Layout_Sheet(SheetName As String)

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    
    ws.Name = SheetName
    
    Worksheets(SheetName).Activate
    
End Function

Function Set_Page_Orientation_xlLandscape(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    Worksheets(SheetName).Activate
    
    With ws.PageSetup
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With

End Function

Function Set_Page_Orientation_xlPortrait(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    Worksheets(SheetName).Activate
    
    With ws.PageSetup
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With

End Function

Function Set_Page_PaperSize(SheetName As String)
Attribute Set_Page_PaperSize.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    Worksheets(SheetName).Activate
    
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4 'xlPaperA3 xlPaperB3
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.text = ""
        .EvenPage.CenterHeader.text = ""
        .EvenPage.RightHeader.text = ""
        .EvenPage.LeftFooter.text = ""
        .EvenPage.CenterFooter.text = ""
        .EvenPage.RightFooter.text = ""
        .FirstPage.LeftHeader.text = ""
        .FirstPage.CenterHeader.text = ""
        .FirstPage.RightHeader.text = ""
        .FirstPage.LeftFooter.text = ""
        .FirstPage.CenterFooter.text = ""
        .FirstPage.RightFooter.text = ""
    End With
    Application.PrintCommunication = True
    
End Function

Function Set_View_xlPageBreakPreview(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    Worksheets(SheetName).Activate
    
    ActiveWindow.View = xlPageBreakPreview

End Function

Function Set_View_xlNormalView(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    Worksheets(SheetName).Activate
    
    ActiveWindow.View = xlNormalView

End Function

Function Set_View_xlPageLayoutView(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    Worksheets(SheetName).Activate
    
    ActiveWindow.View = xlPageLayoutView

End Function

Function Set_Group_Columns(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    Worksheets(SheetName).Activate
    
    ws.Range("B:E").Columns.Group

End Function

Function Set_Group_Rows(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    Worksheets(SheetName).Activate
    
    ws.Range("7:20").Rows.Group

End Function

Function Set_Ungroup_Columns(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    Worksheets(SheetName).Activate
    
    ws.Range("B:E").Columns.Ungroup

End Function

Function Set_Ungroup_Rows(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    Worksheets(SheetName).Activate
    
    ws.Range("7:20").Rows.Ungroup

End Function

Function Set_RemoveDuplicates_No_Header(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    Worksheets(SheetName).Activate
    
    ws.Range("A1") = "AB"
    ws.Range("A2:A5") = "AA"
    ws.Range("A6") = "BB"
    ws.Range("A7:A8") = "aa"
    ws.Range("A9:A14") = "AB"
    ws.Range("A15:A19") = "AA"
    ws.Range("A20:A23") = "bA"
    ws.Range("A24") = "AB"
    ws.Range("B1") = "41"
    ws.Range("B2:B3") = "11"
    ws.Range("B4") = "23"
    ws.Range("B5:B6") = "41"
    ws.Range("B7:B9") = "42"
    ws.Range("B10:B13") = "47"
    ws.Range("B14:B16") = "11"
    ws.Range("B17:B18") = "23"
    ws.Range("B19:B20") = "42"
    ws.Range("B21:B24") = "45"
   
    ws.Range("$A$1:$B$24").RemoveDuplicates Columns:=1, Header:=xlNo
    
End Function

Function Set_RemoveDuplicates_Header(SheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    Worksheets(SheetName).Activate
    
    ws.Range("A1") = "AB"
    ws.Range("A2:A5") = "AA"
    ws.Range("A6") = "BB"
    ws.Range("A7:A8") = "aa"
    ws.Range("A9:A14") = "AB"
    ws.Range("A15:A19") = "AA"
    ws.Range("A20:A23") = "bA"
    ws.Range("A24") = "AB"
    ws.Range("B1") = "41"
    ws.Range("B2:B3") = "11"
    ws.Range("B4") = "23"
    ws.Range("B5:B6") = "41"
    ws.Range("B7:B9") = "42"
    ws.Range("B10:B13") = "47"
    ws.Range("B14:B16") = "11"
    ws.Range("B17:B18") = "23"
    ws.Range("B19:B20") = "42"
    ws.Range("B21:B24") = "45"
   
    ws.Range("$A$1:$B$24").RemoveDuplicates Columns:=2, Header:=xlYes
    
End Function
