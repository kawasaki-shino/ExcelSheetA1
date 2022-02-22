$input = Read-Host "フォルダーを指定してください"
$list = Get-ChildItem -Path $input -Filter *.xlsx -Recurse

$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.EnableEvents = $false

foreach ($file in $list){
    try{
        
        $book = $excel.Workbooks.Open($file.FullName)
        
        foreach ($sheet in $excel.Worksheets){

            if ($sheet.Visible -ne $false){
                $sheet.Activate()
                $sheet.Cells.Item(1,1).Select() | Out-Null
                $excel.ActiveWindow.Zoom = 100
                $excel.ActiveWindow.ScrollColumn = 1
                $excel.ActiveWindow.ScrollRow = 1
            }
        }

        $excel.Worksheets.Item(1).Activate()

        $book.save()
        $book.close()

    } catch {
        Write-Host "Error" -ForegroundColor Red
        echo '$error[0] = ' + $error[0]
    } 
}

$excel.DisplayAlerts = $true
$excel.ScreenUpdating = $true
$excel.EnableEvents = $true
$excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($book) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
[GC]::collect()