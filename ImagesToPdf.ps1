Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$dialog = New-Object System.Windows.Forms.FolderBrowserDialog
if ($dialog.ShowDialog() -ne "OK") { return }

$folder = $dialog.SelectedPath

$images = Get-ChildItem $folder -Include *.png, *.jpg, *.jpeg -Recurse |
          Sort-Object FullName

if ($images.Count -eq 0) { return }

$date = Get-Date -Format "yyyyMMdd"
$outputPdf = Join-Path $folder "$date.pdf"

$printDoc = New-Object System.Drawing.Printing.PrintDocument
$printDoc.PrinterSettings.PrinterName = "Microsoft Print to PDF"
$printDoc.PrinterSettings.PrintToFile = $true
$printDoc.PrinterSettings.PrintFileName = $outputPdf

$script:current = 0   # ← 重要（scriptスコープ）

$printDoc.add_PrintPage({
    param($sender, $e)

    if ($script:current -ge $images.Count) {
        $e.HasMorePages = $false
        return
    }

    $img = [System.Drawing.Image]::FromFile($images[$script:current].FullName)

    $ratioX = $e.MarginBounds.Width / $img.Width
    $ratioY = $e.MarginBounds.Height / $img.Height
    $ratio = [Math]::Min($ratioX, $ratioY)

    $width  = $img.Width * $ratio
    $height = $img.Height * $ratio

    $x = ($e.MarginBounds.Width - $width) / 2
    $y = ($e.MarginBounds.Height - $height) / 2

    $e.Graphics.DrawImage($img, $x, $y, $width, $height)

    $img.Dispose()

    $script:current++

    $e.HasMorePages = ($script:current -lt $images.Count)
})

$printDoc.Print()

Write-Host "完了: $outputPdf"