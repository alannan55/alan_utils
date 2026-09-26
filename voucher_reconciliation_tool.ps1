param(
    [string]$WorkbookPath = "",
    [string]$OutputDirectory = "",
    [switch]$RunNow
)

Set-StrictMode -Version 2.0

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$Script:XlScreen = 1
$Script:XlBitmap = 2
$Script:XlSolid = 1
$Script:TempSheetName = "__voucher_shot_tmp__"

function New-UnicodeString {
    param([int[]]$CodePoints)
    $chars = foreach ($codePoint in $CodePoints) {
        [char]$codePoint
    }
    return -join $chars
}

$Script:TotalText = New-UnicodeString -CodePoints @(0x603B, 0x8BA1)
$Script:FullLeftParen = [string][char]0xFF08
$Script:FullRightParen = [string][char]0xFF09

$specialNames = @(
    (New-UnicodeString -CodePoints @(0x5F90, 0x6625, 0x6770)),
    (New-UnicodeString -CodePoints @(0x82CF, 0x6653, 0x60A6)),
    (New-UnicodeString -CodePoints @(0x90ED, 0xFF08, 0x6708, 0x5E95, 0x7FA4, 0x91CC, 0x53D1, 0xFF09)),
    (New-UnicodeString -CodePoints @(0x5065, 0x5174, 0x5F20, 0x8F89))
)
$Script:SpecialNameSet = @{}
foreach ($name in $specialNames) {
    $Script:SpecialNameSet[$name] = $true
}

function Convert-HexToExcelColor {
    param([Parameter(Mandatory = $true)][string]$HexColor)

    $hex = $HexColor.Trim().TrimStart("#")
    if ($hex.Length -ne 6) {
        throw "Invalid color: $HexColor"
    }

    $red = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $green = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $blue = [Convert]::ToInt32($hex.Substring(4, 2), 16)
    return $red + ($green * 256) + ($blue * 65536)
}

$Script:GreenColor = Convert-HexToExcelColor "#92D050"
$Script:RedColor = Convert-HexToExcelColor "#FF0000"

function Normalize-CellText {
    param($Value)

    if ($null -eq $Value) {
        return ""
    }

    $text = ([string]$Value).Trim()
    $text = $text.Replace("(", $Script:FullLeftParen).Replace(")", $Script:FullRightParen)
    $text = $text -replace "\s+", ""
    return $text
}

function Convert-CellNumber {
    param($Value)

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal] -or $Value -is [single]) {
        return [double]$Value
    }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $text = $text.Replace(",", "").Replace([string][char]0xFFE5, "").Replace([string][char]0x00A5, "")
    $styles = [System.Globalization.NumberStyles]::Float -bor [System.Globalization.NumberStyles]::AllowThousands

    $number = 0.0
    if ([double]::TryParse($text, $styles, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $number
    }
    if ([double]::TryParse($text, $styles, [System.Globalization.CultureInfo]::CurrentCulture, [ref]$number)) {
        return $number
    }

    return $null
}

function Test-CellZero {
    param($Value)

    $number = Convert-CellNumber $Value
    return ($null -ne $number -and [Math]::Abs($number) -lt 0.0000001)
}

function Set-CellFill {
    param($Cell, [int]$Color)

    $Cell.Interior.Pattern = $Script:XlSolid
    $Cell.Interior.Color = $Color
}

function Get-LastUsedRow {
    param($Worksheet)

    $usedRange = $Worksheet.UsedRange
    return [int]($usedRange.Row + $usedRange.Rows.Count - 1)
}

function Get-SafeFileNamePart {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return "blank"
    }

    $safe = $Text.Trim()
    foreach ($invalidChar in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace([string]$invalidChar, "_")
    }
    $safe = $safe -replace "\s+", "_"
    if ($safe.Length -gt 80) {
        $safe = $safe.Substring(0, 80)
    }
    return $safe
}

function Add-LogLine {
    param($TextBox, [string]$Message)

    $line = "[{0}] {1}" -f (Get-Date).ToString("HH:mm:ss"), $Message
    $TextBox.AppendText($line + [Environment]::NewLine)
    $TextBox.SelectionStart = $TextBox.TextLength
    $TextBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Find-NextTotalRow {
    param(
        [System.Collections.Generic.List[int]]$TotalRows,
        [int]$AfterRow
    )

    foreach ($row in $TotalRows) {
        if ($row -gt $AfterRow) {
            return $row
        }
    }
    return $null
}

function Process-Worksheet {
    param(
        $Worksheet,
        [int]$SheetIndex,
        [scriptblock]$Log
    )

    $lastRow = Get-LastUsedRow $Worksheet
    $totalRows = New-Object "System.Collections.Generic.List[int]"
    $specialRows = New-Object "System.Collections.Generic.List[int]"

    for ($row = 1; $row -le $lastRow; $row++) {
        $gText = Normalize-CellText $Worksheet.Cells.Item($row, 7).Value2
        if ($gText -eq $Script:TotalText) {
            $totalRows.Add([int]$row)
        }
        elseif ($Script:SpecialNameSet.ContainsKey($gText)) {
            $specialRows.Add([int]$row)
        }
    }

    $specialTargetRows = @{}
    foreach ($specialRow in $specialRows) {
        $targetRow = Find-NextTotalRow -TotalRows $totalRows -AfterRow $specialRow
        if ($null -ne $targetRow) {
            $specialTargetRows[[int]$targetRow] = $true
        }
        else {
            & $Log ("Sheet {0}: special row {1} has no following total row." -f $Worksheet.Name, $specialRow)
        }
    }

    $pendingTasks = New-Object "System.Collections.Generic.List[object]"
    $greenCount = 0
    $redCount = 0

    for ($index = 0; $index -lt $totalRows.Count; $index++) {
        $totalRow = [int]$totalRows[$index]
        $amountCell = $Worksheet.Cells.Item($totalRow, 10)
        $isZero = Test-CellZero $amountCell.Value2
        $isSpecialTarget = $specialTargetRows.ContainsKey($totalRow)

        if ($isZero -or $isSpecialTarget) {
            Set-CellFill -Cell $amountCell -Color $Script:GreenColor
            $greenCount++
        }
        else {
            Set-CellFill -Cell $amountCell -Color $Script:RedColor
            $redCount++

            if ($index -eq 0) {
                $startRow = 1
            }
            else {
                $startRow = [int]$totalRows[$index - 1] + 1
            }

            $pendingTasks.Add([pscustomobject]@{
                SheetIndex = $SheetIndex
                SheetName = [string]$Worksheet.Name
                StartRow = [int]$startRow
                EndRow = [int]$totalRow
                AmountText = [string]$amountCell.Text
            })
        }
    }

    & $Log ("Sheet {0}: totals={1}, green={2}, red/pending={3}." -f $Worksheet.Name, $totalRows.Count, $greenCount, $redCount)
    return $pendingTasks
}

function Remove-TempSheet {
    param($Workbook)

    for ($index = $Workbook.Worksheets.Count; $index -ge 1; $index--) {
        $sheet = $Workbook.Worksheets.Item($index)
        if ($sheet.Name -eq $Script:TempSheetName) {
            $sheet.Delete() | Out-Null
        }
    }
}

function Reset-TempSheet {
    param($TempWorksheet)

    try {
        $TempWorksheet.Cells.UnMerge() | Out-Null
    }
    catch {
        # The sheet may not have merged cells yet.
    }

    $TempWorksheet.Cells.Clear() | Out-Null
    while ($TempWorksheet.Shapes.Count -gt 0) {
        $TempWorksheet.Shapes.Item(1).Delete() | Out-Null
    }
}

function Build-TempScreenshotRange {
    param(
        $SourceWorksheet,
        $TempWorksheet,
        [int]$StartRow,
        [int]$EndRow
    )

    Reset-TempSheet $TempWorksheet | Out-Null

    for ($column = 1; $column -le 10; $column++) {
        $TempWorksheet.Columns.Item($column).ColumnWidth = $SourceWorksheet.Columns.Item($column).ColumnWidth
    }

    if ($StartRow -le 2) {
        $rowCount = $EndRow
        $SourceWorksheet.Range($SourceWorksheet.Cells.Item(1, 1), $SourceWorksheet.Cells.Item($EndRow, 10)).Copy($TempWorksheet.Cells.Item(1, 1)) | Out-Null
        for ($row = 1; $row -le $EndRow; $row++) {
            $TempWorksheet.Rows.Item($row).RowHeight = $SourceWorksheet.Rows.Item($row).RowHeight
        }
    }
    else {
        $SourceWorksheet.Range($SourceWorksheet.Cells.Item(1, 1), $SourceWorksheet.Cells.Item(2, 10)).Copy($TempWorksheet.Cells.Item(1, 1)) | Out-Null
        $SourceWorksheet.Range($SourceWorksheet.Cells.Item($StartRow, 1), $SourceWorksheet.Cells.Item($EndRow, 10)).Copy($TempWorksheet.Cells.Item(3, 1)) | Out-Null

        $TempWorksheet.Rows.Item(1).RowHeight = $SourceWorksheet.Rows.Item(1).RowHeight
        $TempWorksheet.Rows.Item(2).RowHeight = $SourceWorksheet.Rows.Item(2).RowHeight

        $targetRow = 3
        for ($sourceRow = $StartRow; $sourceRow -le $EndRow; $sourceRow++) {
            $TempWorksheet.Rows.Item($targetRow).RowHeight = $SourceWorksheet.Rows.Item($sourceRow).RowHeight
            $targetRow++
        }
        $rowCount = 2 + ($EndRow - $StartRow + 1)
    }

    return ,($TempWorksheet.Range($TempWorksheet.Cells.Item(1, 1), $TempWorksheet.Cells.Item($rowCount, 10)))
}

function Save-RangeAsPng {
    param(
        $Range,
        [string]$OutputPath
    )

    try {
        [System.Windows.Forms.Clipboard]::Clear()
    }
    catch {
        # Another process can temporarily own the clipboard; retry below.
    }

    for ($attempt = 1; $attempt -le 6; $attempt++) {
        try {
            $Range.Worksheet.Activate() | Out-Null
            $Range.Select() | Out-Null
        }
        catch {
            # CopyPicture can still work without selection in many Excel builds.
        }
        $Range.CopyPicture($Script:XlScreen, $Script:XlBitmap) | Out-Null
        Start-Sleep -Milliseconds (200 * $attempt)

        $image = [System.Windows.Forms.Clipboard]::GetImage()
        if ($null -ne $image) {
            try {
                if (Test-Path -LiteralPath $OutputPath) {
                    Remove-Item -LiteralPath $OutputPath -Force
                }
                $image.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
            }
            finally {
                $image.Dispose()
            }

            try {
                [System.Windows.Forms.Clipboard]::Clear()
            }
            catch {
            }
            return
        }
    }

    throw "Excel did not provide an image on the clipboard. Please try again after closing clipboard/screenshot tools."
}

function Export-PendingScreenshots {
    param(
        $Workbook,
        [System.Collections.Generic.List[object]]$Tasks,
        [string]$OutputDirectory,
        [scriptblock]$Log
    )

    if ($Tasks.Count -eq 0) {
        & $Log "No pending rows found, so no screenshots are needed."
        return 0
    }

    $Workbook.Application.ScreenUpdating = $true
    Remove-TempSheet $Workbook
    $tempWorksheet = $Workbook.Worksheets.Add([System.Type]::Missing, $Workbook.Worksheets.Item($Workbook.Worksheets.Count))
    $tempWorksheet.Name = $Script:TempSheetName

    try {
        $counter = 1
        foreach ($task in $Tasks) {
            $sourceWorksheet = $Workbook.Worksheets.Item($task.SheetIndex)
            $range = Build-TempScreenshotRange -SourceWorksheet $sourceWorksheet -TempWorksheet $tempWorksheet -StartRow $task.StartRow -EndRow $task.EndRow

            $sheetPart = Get-SafeFileNamePart $task.SheetName
            $amountPart = Get-SafeFileNamePart $task.AmountText
            $fileName = "{0:D3}_{1}_rows_{2}-{3}_amount_{4}.png" -f $counter, $sheetPart, $task.StartRow, $task.EndRow, $amountPart
            $outputPath = Join-Path $OutputDirectory $fileName

            Save-RangeAsPng -Range $range -OutputPath $outputPath
            & $Log ("Screenshot {0}/{1}: {2}" -f $counter, $Tasks.Count, $fileName)
            $counter++
        }

        return $Tasks.Count
    }
    finally {
        $tempWorksheet.Delete() | Out-Null
    }
}

function Invoke-VoucherReconciliation {
    param(
        [Parameter(Mandatory = $true)][string]$WorkbookPath,
        [Parameter(Mandatory = $true)][string]$OutputDirectory,
        [scriptblock]$Log
    )

    if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) {
        throw "Workbook not found: $WorkbookPath"
    }

    if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
    }

    $excel = $null
    $workbook = $null

    try {
        & $Log "Starting Excel..."
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false
        $excel.EnableEvents = $false

        & $Log "Opening workbook..."
        $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $false)
        if ($workbook.ReadOnly) {
            throw "The workbook opened as read-only. Please close it in Excel first, then run this tool again."
        }

        if ($workbook.Worksheets.Count -lt 2) {
            throw "The workbook has fewer than two worksheets."
        }

        $allTasks = New-Object "System.Collections.Generic.List[object]"
        for ($sheetIndex = 1; $sheetIndex -le 2; $sheetIndex++) {
            $worksheet = $workbook.Worksheets.Item($sheetIndex)
            & $Log ("Processing sheet {0}: {1}" -f $sheetIndex, $worksheet.Name)
            $tasks = Process-Worksheet -Worksheet $worksheet -SheetIndex $sheetIndex -Log $Log
            foreach ($task in $tasks) {
                $allTasks.Add($task)
            }
        }

        & $Log ("Pending screenshot tasks: {0}" -f $allTasks.Count)
        $screenshotCount = Export-PendingScreenshots -Workbook $workbook -Tasks $allTasks -OutputDirectory $OutputDirectory -Log $Log

        Remove-TempSheet $workbook
        & $Log "Saving workbook in place..."
        $workbook.Save()
        $workbook.Close($true)
        $workbook = $null

        & $Log ("Done. Screenshots exported: {0}" -f $screenshotCount)
    }
    catch {
        if ($null -ne $workbook) {
            $workbook.Close($false)
        }
        throw
    }
    finally {
        if ($null -ne $excel) {
            $excel.Quit()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function New-LabeledTextBox {
    param(
        $Form,
        [string]$LabelText,
        [int]$Top
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $LabelText
    $label.Left = 14
    $label.Top = $Top + 4
    $label.Width = 130
    $Form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Left = 145
    $textBox.Top = $Top
    $textBox.Width = 500
    $Form.Controls.Add($textBox)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Browse..."
    $button.Left = 655
    $button.Top = $Top - 1
    $button.Width = 85
    $button.Height = 25
    $Form.Controls.Add($button)

    return [pscustomobject]@{
        TextBox = $textBox
        Button = $button
    }
}

function Show-MainForm {
    if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne "STA") {
        [System.Windows.Forms.MessageBox]::Show(
            "Clipboard screenshots require STA mode. Please start this tool with run_voucher_reconciliation_tool.bat.",
            "Startup mode",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Voucher Reconciliation Screenshot Tool"
    $form.StartPosition = "CenterScreen"
    $form.Width = 775
    $form.Height = 455
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false

    $workbookInput = New-LabeledTextBox -Form $form -LabelText "Workbook path" -Top 22
    $outputInput = New-LabeledTextBox -Form $form -LabelText "Screenshot folder" -Top 62

    $workbookInput.Button.Add_Click({
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = "Excel files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|All files (*.*)|*.*"
        $dialog.Title = "Select workbook"
        if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
            $workbookInput.TextBox.Text = $dialog.FileName
        }
    })

    $outputInput.Button.Add_Click({
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = "Select screenshot folder"
        if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
            $outputInput.TextBox.Text = $dialog.SelectedPath
        }
    })

    $runButton = New-Object System.Windows.Forms.Button
    $runButton.Text = "Run"
    $runButton.Left = 655
    $runButton.Top = 102
    $runButton.Width = 85
    $runButton.Height = 30
    $form.Controls.Add($runButton)

    $logBox = New-Object System.Windows.Forms.TextBox
    $logBox.Left = 14
    $logBox.Top = 148
    $logBox.Width = 726
    $logBox.Height = 250
    $logBox.Multiline = $true
    $logBox.ScrollBars = "Vertical"
    $logBox.ReadOnly = $true
    $form.Controls.Add($logBox)

    $runButton.Add_Click({
        $workbookPath = $workbookInput.TextBox.Text.Trim()
        $outputDirectory = $outputInput.TextBox.Text.Trim()

        if ([string]::IsNullOrWhiteSpace($workbookPath) -or [string]::IsNullOrWhiteSpace($outputDirectory)) {
            [System.Windows.Forms.MessageBox]::Show($form, "Please select both paths.", "Missing path", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            $form,
            "This will modify the selected workbook in place and save it. Use a copy if you are testing.`r`n`r`nContinue?",
            "Confirm in-place save",
            [System.Windows.Forms.MessageBoxButtons]::OKCancel,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::OK) {
            return
        }

        $runButton.Enabled = $false
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $logBox.Clear()

        try {
            Invoke-VoucherReconciliation -WorkbookPath $workbookPath -OutputDirectory $outputDirectory -Log {
                param([string]$message)
                Add-LogLine -TextBox $logBox -Message $message
            }
            [System.Windows.Forms.MessageBox]::Show($form, "Finished.", "Done", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
        catch {
            Add-LogLine -TextBox $logBox -Message ("ERROR: " + $_.Exception.Message)
            [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
        finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $runButton.Enabled = $true
        }
    })

    [void]$form.ShowDialog()
}

if ($RunNow) {
    Invoke-VoucherReconciliation -WorkbookPath $WorkbookPath -OutputDirectory $OutputDirectory -Log {
        param([string]$message)
        Write-Host ("[{0}] {1}" -f (Get-Date).ToString("HH:mm:ss"), $message)
    }
}
else {
    Show-MainForm
}
