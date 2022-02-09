$ErrorActionPreference = "Stop" #例外が出たらその時点で終了
$srcDir = (Resolve-Path $args[0]).Path
$dstDir = (Resolve-Path $args[1]).Path
try{

    #ExcelObject
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Get-ChildItem -Path $srcDir -Filter "*.xls" | % {
        $dstPath = Join-Path $dstPath $($_.BaseName + ".xlsx")
        if(-not (Test-Path -Path $dstPath)) {
            $book = $excel.Workbook.Open($_.FullName)
            $book.SaveAs($dstPath, 51)
            $book.Close()
        }
        else {

        }
    }
} finally {
    $excel.Quit()
}