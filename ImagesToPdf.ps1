Add-Type -AssemblyName System.Drawing

# 出力ファイル名のサフィックス（例: _領収書）。必要に応じて編集してください。
$FilenameSuffix = "_領収書"

# フォルダはこのスクリプトのあるフォルダを使用
$folder = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
if (-not $folder) { $folder = Get-Location }

$images = Get-ChildItem -Path $folder -File |
          Where-Object { $_.Extension -match '(?i)\.(png|jpe?g)$' } |
          Sort-Object FullName

if ($images.Count -eq 0) {
    Write-Host "画像ファイルが見つかりません: $folder"
    return
}

$date = Get-Date -Format "yyyyMMdd"
$outputPdf = Join-Path $folder "$date$FilenameSuffix.pdf"

if (Test-Path $outputPdf) {
    $answer = Read-Host "出力ファイル '$outputPdf' は既に存在します。上書きしますか？ (y/n)"
    if ($answer -notin @('y','Y','yes','Yes','はい')) {
        Write-Host "処理を中止しました"
        Read-Host "Enterキーで終了します"
        return
    }
    try {
        Remove-Item -Path $outputPdf -Force -ErrorAction Stop
    } catch {
        Write-Host "上書き準備に失敗しました（ファイルが開かれている可能性があります）： $outputPdf"
        return
    }
}

$printDoc = New-Object System.Drawing.Printing.PrintDocument
$printDoc.PrinterSettings.PrinterName = "Microsoft Print to PDF"
$printDoc.PrinterSettings.PrintToFile = $true
$printDoc.PrinterSettings.PrintFileName = $outputPdf

$script:current = 0

$printDoc.add_PrintPage({
    param($sender, $e)

    if ($script:current -ge $images.Count) {
        $e.HasMorePages = $false
        return
    }

    $img = [System.Drawing.Image]::FromFile($images[$script:current].FullName)

    # EXIF Orientationを適用
    $orientationId = 0x0112
    if ($img.PropertyIdList -contains $orientationId) {
        $prop = $img.GetPropertyItem($orientationId)
        $orientation = $prop.Value[0]

        switch ($orientation) {
            3 { $img.RotateFlip([System.Drawing.RotateFlipType]::Rotate180FlipNone) }
            6 { $img.RotateFlip([System.Drawing.RotateFlipType]::Rotate90FlipNone) }
            8 { $img.RotateFlip([System.Drawing.RotateFlipType]::Rotate270FlipNone) }
        }

        # Orientationタグ削除（再回転防止）
        $img.RemovePropertyItem($orientationId)
    }

    $ratioX = $e.MarginBounds.Width / $img.Width
    $ratioY = $e.MarginBounds.Height / $img.Height
    $ratio = [Math]::Min($ratioX, $ratioY)

    $width  = $img.Width * $ratio
    $height = $img.Height * $ratio

    $x = $e.MarginBounds.Left + ($e.MarginBounds.Width - $width) / 2
    $y = $e.MarginBounds.Top  + ($e.MarginBounds.Height - $height) / 2

    $e.Graphics.DrawImage($img, $x, $y, $width, $height)

    $img.Dispose()

    $script:current++

    $e.HasMorePages = ($script:current -lt $images.Count)
})

try {
    $printDoc.Print()
} catch {
    Write-Host "印刷中にエラーが発生しました。出力をスキップします: $outputPdf"
    Write-Host $_.Exception.Message
    return
}

Write-Host "完了: $outputPdf"