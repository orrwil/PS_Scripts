

## Create a spreadsheet with test data

$excel = New-Object -ComObject Excel.Application
$excel.visible = $true
$workbook = $excel.Workbooks.Add()
$excel.Cells.item(1,1) = "Cell 1,1"
$workbook.SaveAs("C:\Scripts\Test.xlsx")
$excel.Quit()

## Reoprn newly created sheet
$excel = New-Object -ComObject Excel.Application
$excel.visible = $true
$workbook = $excel.Workbooks.Open("C:\Scripts\Test.xlsx")
