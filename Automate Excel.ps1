# Imports
. ".\PS functions\Excel Functions.ps1"


# Paths
$currentDir = (Get-Item -Path ".\").FullName
$workbookName = "Input.xlsx"
$workbookPath = "{0}\{1}" -F $currentDir, $workbookName

$workbook = OpenExcelBook -FileName $workbookPath
$worksheet =  $workbook.sheets.item("Sheet1")

$worksheet.Cells.Item(1, 1) = "ahh"
$A1 = $worksheet.Range("A1").Text
echo $A1

$workbook.save()
$workbook.close()
