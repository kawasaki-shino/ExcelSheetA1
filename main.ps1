# 対象のExcelを取得
$path = Read-Host "フォルダーまたはExcelファイルを指定してください"
$list = Get-ChildItem -Path $path -Filter *.xlsx -Recurse -File
$total = ($list | ? { ! $_.PsIsContainer }).Count

if ($total -eq 0){
    Write-Host '指定されたフォルダーに「.xlsx」のファイルが見つからないか、指定されたファイルが「.xlsx」ではありません。処理を中断します。'
    exit
}

# Excel開いて、高速化設定
$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.EnableEvents = $false


# ファイル処理
foreach ($file in $list){

    # ブックを開く
    $book = $excel.Workbooks.Open($file.FullName)

    # 進捗表示
    $counter ++;
    $status = $counter.ToString() + "/" + $total.ToString()
    $per = $counter / $total * 100
            
    Write-Progress -Activity "進捗" -Status $status -PercentComplete $per
            
    # 現在開いてるブックのシートに対してフォーカスをA1に移動させる等の処理をする
    foreach ($sheet in $excel.Worksheets){

        if ($sheet.Visible -ne $false){
            $sheet.Activate()
            $sheet.Cells.Item(1,1).Select() | Out-Null
            $excel.ActiveWindow.Zoom = 100
            $excel.ActiveWindow.ScrollColumn = 1
            $excel.ActiveWindow.ScrollRow = 1
        }
    }

    # シートの一枚目にフォーカスを移動させる
    $excel.Worksheets.Item(1).Activate()

    # 保存
    $book.save()
        
    # ブックを閉じる
    $book.close()
}

# Excelの設定を戻して閉じる
$excel.DisplayAlerts = $true
$excel.ScreenUpdating = $true
$excel.EnableEvents = $true
$excel.Quit()

# 変数開放
Remove-Variable sheet
Remove-Variable book
Remove-Variable excel
Remove-Variable file
Remove-Variable path
Remove-Variable list
Remove-Variable counter
Remove-Variable total
Remove-Variable status
Remove-Variable per

[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[System.GC]::Collect()

1|%{$_} > $null
[System.GC]::Collect()

