Function OpenExcelBook($FileName) {
    $Excel=new-object -ComObject Excel.Application
    Return $Excel.workbooks.open($Filename)
}

Function SaveExcelBook($Workbook) {
    $Workbook.save()
    $Workbook.close()
}

Function ReadCellData($WorkSheet,$Cell) {
    Return $WorkSheet.Range($Cell).Text
}