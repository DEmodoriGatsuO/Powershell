$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
Get-ChildItem -Filter "*.xls" | % {
    $xlsxFile = $_.FullName.Replace(".xls",".xlsx")
    $book = $excel.Workbooks.Open($_.FullName)
    $book.SaveAs($xlsxFile, 51)
    $book.Close()
}