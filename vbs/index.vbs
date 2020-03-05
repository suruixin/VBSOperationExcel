option explicit
Dim oExcel, xlmodule, strCode, fE
Set oExcel = CreateObject( "Excel.Application" )
oExcel.Visible = True
oExcel.DisplayAlerts = False
set fE = oExcel.WorkBooks.Open("G:\\vbs\\demo\\excel\\demo.xls")
Set xlmodule = fE.VBProject.VBComponents.Add(1) 
strCode = _ 
	"Sub print_excel()" & vbCr & _ 
    "Application.PrintCommunication = False" & vbCr & _ 
    "With ActiveSheet.PageSetup" & vbCr & _ 
    "    .PrintTitleRows = """ & vbCr & _ 
    "    .PrintTitleColumns = """ & vbCr & _ 
    "End With" & vbCr & _ 
    "Application.PrintCommunication = True" & vbCr & _ 
    "ActiveSheet.PageSetup.PrintArea = """ & vbCr & _ 
    "Application.PrintCommunication = False" & vbCr & _ 
    "With ActiveSheet.PageSetup" & vbCr & _ 
    "    .LeftHeader = """ & vbCr & _ 
    "    .CenterHeader = """ & vbCr & _ 
    "    .RightHeader = """ & vbCr & _ 
    "    .LeftFooter = """ & vbCr & _ 
    "    .CenterFooter = """ & vbCr & _ 
    "    .RightFooter = """ & vbCr & _ 
    "    .LeftMargin = Application.InchesToPoints(0.708661417322835)" & vbCr & _ 
    "    .RightMargin = Application.InchesToPoints(0.708661417322835)" & vbCr & _ 
    "    .TopMargin = Application.InchesToPoints(0.748031496062992)" & vbCr & _ 
    "    .BottomMargin = Application.InchesToPoints(0.748031496062992)" & vbCr & _ 
    "    .HeaderMargin = Application.InchesToPoints(0.31496062992126)" & vbCr & _ 
    "    .FooterMargin = Application.InchesToPoints(0.31496062992126)" & vbCr & _ 
    "    .PrintHeadings = False" & vbCr & _ 
    "    .PrintGridlines = False" & vbCr & _ 
    "    .PrintComments = xlPrintNoComments" & vbCr & _ 
    "    .PrintQuality = 600" & vbCr & _ 
    "    .CenterHorizontally = True" & vbCr & _ 
    "    .CenterVertically = True" & vbCr & _ 
    "    .Orientation = xlPortrait" & vbCr & _ 
    "    .Draft = False" & vbCr & _ 
    "    .PaperSize = xlPaperA4" & vbCr & _ 
    "    .FirstPageNumber = xlAutomatic" & vbCr & _ 
    "    .Order = xlDownThenOver" & vbCr & _ 
    "    .BlackAndWhite = False" & vbCr & _ 
    "    .Zoom = 100" & vbCr & _ 
    "    .PrintErrors = xlPrintErrorsDisplayed" & vbCr & _ 
    "    .OddAndEvenPagesHeaderFooter = False" & vbCr & _ 
    "    .DifferentFirstPageHeaderFooter = False" & vbCr & _ 
    "    .ScaleWithDocHeaderFooter = True" & vbCr & _ 
    "    .AlignMarginsHeaderFooter = True" & vbCr & _ 
    "    .EvenPage.LeftHeader.Text = """ & vbCr & _ 
    "    .EvenPage.CenterHeader.Text = """ & vbCr & _ 
    "    .EvenPage.RightHeader.Text = """ & vbCr & _ 
    "    .EvenPage.LeftFooter.Text = """ & vbCr & _ 
    "    .EvenPage.CenterFooter.Text = """ & vbCr & _ 
    "    .EvenPage.RightFooter.Text = """ & vbCr & _ 
    "    .FirstPage.LeftHeader.Text = """ & vbCr & _ 
    "    .FirstPage.CenterHeader.Text = """ & vbCr & _ 
    "    .FirstPage.RightHeader.Text = """ & vbCr & _ 
    "    .FirstPage.LeftFooter.Text = """ & vbCr & _ 
    "    .FirstPage.CenterFooter.Text = """ & vbCr & _ 
    "    .FirstPage.RightFooter.Text = """ & vbCr & _ 
    "End With" & vbCr & _ 
    "Application.PrintCommunication = True" & vbCr & _ 
    "ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _" & vbCr & _ 
    "    IgnorePrintAreas:=False" & vbCr & _ 
	"End Sub"

 xlmodule.CodeModule.AddFromString strCode
 fE.Application.Run "print_excel"
 fE.Close
 oExcel.Quit
Set oExcel= nothing
Set fE= nothing
 