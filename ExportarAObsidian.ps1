Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[Windows.Forms.Application]::EnableVisualStyles()

$script:wordApplication = $null
$script:excelApplication = $null
$script:powerPointApplication = $null
$script:logPath = $null
$script:counts = @{
    Files       = 0
    Emails      = 0
    Attachments = 0
    Skipped     = 0
    Errors      = 0
}

function Write-ExportLog {
    param([string]$Message)

    $line = '{0}  {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    $logBox.AppendText($line + [Environment]::NewLine)
    $logBox.SelectionStart = $logBox.TextLength
    $logBox.ScrollToCaret()
    [Windows.Forms.Application]::DoEvents()
    if ($script:logPath) {
        [IO.File]::AppendAllText($script:logPath, $line + [Environment]::NewLine, [Text.Encoding]::UTF8)
    }
}

function Get-SafeName {
    param([string]$Name)

    $safe = $Name
    foreach ($character in [IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace([string]$character, '_')
    }
    $safe = $safe.Trim().TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($safe)) {
        return 'Sin nombre'
    }
    return $safe
}

function Get-UniquePath {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return $Path
    }
    $directory = [IO.Path]::GetDirectoryName($Path)
    $name = [IO.Path]::GetFileNameWithoutExtension($Path)
    $extension = [IO.Path]::GetExtension($Path)
    $index = 2
    do {
        $candidate = Join-Path $directory ('{0} ({1}){2}' -f $name, $index, $extension)
        $index++
    } while (Test-Path -LiteralPath $candidate)
    return $candidate
}

function ConvertTo-MarkdownScalar {
    param([AllowNull()][object]$Value)
    return ConvertTo-Json -InputObject ([string]$Value) -Compress
}

function ConvertTo-MarkdownLink {
    param(
        [string]$FromFile,
        [string]$ToFile,
        [string]$Label
    )

    $fromUri = New-Object Uri(($FromFile -replace '\\', '/') + '/')
    $toUri = New-Object Uri($ToFile -replace '\\', '/')
    $relative = [Uri]::UnescapeDataString($fromUri.MakeRelativeUri($toUri).ToString())
    $encoded = (($relative -split '/') | ForEach-Object { [Uri]::EscapeDataString($_) }) -join '/'
    $escapedLabel = $Label.Replace('\', '\\').Replace('[', '\[').Replace(']', '\]')
    return ('[{0}]({1})' -f $escapedLabel, $encoded)
}

function Get-WordText {
    param([string]$Path)

    if (-not $script:wordApplication) {
        $script:wordApplication = New-Object -ComObject Word.Application
        $script:wordApplication.Visible = $false
        $script:wordApplication.DisplayAlerts = 0
    }
    $document = $script:wordApplication.Documents.Open($Path, $false, $true, $false)
    try {
        return [string]$document.Content.Text
    }
    finally {
        $document.Close(0)
    }
}

function ConvertTo-MarkdownTable {
    param([object[]]$Rows)

    if (-not $Rows -or $Rows.Count -eq 0) {
        return ''
    }
    $headers = @($Rows[0].PSObject.Properties.Name)
    if ($headers.Count -eq 0) {
        return ''
    }
    $lines = @(
        '| ' + (($headers | ForEach-Object { ([string]$_).Replace('|', '\|') }) -join ' | ') + ' |'
        '| ' + (($headers | ForEach-Object { '---' }) -join ' | ') + ' |'
    )
    foreach ($row in $Rows) {
        $values = foreach ($header in $headers) {
            ([string]$row.$header).Replace('|', '\|').Replace("`r", ' ').Replace("`n", ' ')
        }
        $lines += '| ' + ($values -join ' | ') + ' |'
    }
    return $lines -join [Environment]::NewLine
}

function Get-ExcelMarkdown {
    param([string]$Path)

    if (-not $script:excelApplication) {
        $script:excelApplication = New-Object -ComObject Excel.Application
        $script:excelApplication.Visible = $false
        $script:excelApplication.DisplayAlerts = $false
    }
    $workbook = $script:excelApplication.Workbooks.Open($Path, 0, $true)
    try {
        $sections = @()
        foreach ($sheet in $workbook.Worksheets) {
            $range = $sheet.UsedRange
            $rowCount = [Math]::Min([int]$range.Rows.Count, 300)
            $columnCount = [Math]::Min([int]$range.Columns.Count, 30)
            $lines = @('## ' + $sheet.Name, '')
            if ($rowCount -gt 0 -and $columnCount -gt 0) {
                for ($row = 1; $row -le $rowCount; $row++) {
                    $cells = @()
                    for ($column = 1; $column -le $columnCount; $column++) {
                        $value = [string]$range.Cells.Item($row, $column).Text
                        $cells += $value.Replace('|', '\|').Replace("`r", ' ').Replace("`n", ' ')
                    }
                    $lines += '| ' + ($cells -join ' | ') + ' |'
                    if ($row -eq 1) {
                        $lines += '| ' + ((1..$columnCount | ForEach-Object { '---' }) -join ' | ') + ' |'
                    }
                }
            }
            if ([int]$range.Rows.Count -gt $rowCount -or [int]$range.Columns.Count -gt $columnCount) {
                $lines += ''
                $lines += '> Tabla limitada a 300 filas y 30 columnas.'
            }
            $sections += ($lines -join [Environment]::NewLine)
        }
        return $sections -join "`r`n`r`n"
    }
    finally {
        $workbook.Close($false)
    }
}

function Get-PowerPointMarkdown {
    param([string]$Path)

    if (-not $script:powerPointApplication) {
        $script:powerPointApplication = New-Object -ComObject PowerPoint.Application
    }
    $presentation = $script:powerPointApplication.Presentations.Open($Path, $true, $false, $false)
    try {
        $sections = @()
        foreach ($slide in $presentation.Slides) {
            $lines = @('## Diapositiva ' + $slide.SlideIndex, '')
            foreach ($shape in $slide.Shapes) {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    $text = ([string]$shape.TextFrame.TextRange.Text).Trim()
                    if ($text) {
                        $lines += $text
                        $lines += ''
                    }
                }
            }
            $sections += ($lines -join [Environment]::NewLine)
        }
        return $sections -join "`r`n"
    }
    finally {
        $presentation.Close()
    }
}

function Convert-FileToMarkdown {
    param([string]$Path)

    $extension = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    switch ($extension) {
        { $_ -in '.txt', '.md', '.markdown', '.json', '.xml', '.log' } {
            return [IO.File]::ReadAllText($Path)
        }
        '.csv' {
            $rows = @(Import-Csv -LiteralPath $Path)
            if ($rows.Count -eq 0) {
                return ''
            }
            return ConvertTo-MarkdownTable -Rows $rows
        }
        { $_ -in '.html', '.htm' } {
            $html = [IO.File]::ReadAllText($Path)
            $html = [regex]::Replace($html, '(?is)<(script|style)\b[^>]*>.*?</\1>', '')
            $html = [regex]::Replace($html, '(?i)<\s*(br|/p|/div|/li|/h[1-6]|/tr)\b[^>]*>', "`r`n")
            $text = [regex]::Replace($html, '<[^>]+>', '')
            return [Net.WebUtility]::HtmlDecode($text).Trim()
        }
        { $_ -in '.docx', '.doc', '.rtf', '.pdf' } {
            return Get-WordText -Path $Path
        }
        { $_ -in '.xlsx', '.xls', '.xlsm' } {
            return Get-ExcelMarkdown -Path $Path
        }
        { $_ -in '.pptx', '.ppt' } {
            return Get-PowerPointMarkdown -Path $Path
        }
        default {
            return $null
        }
    }
}

function Export-OneDriveFiles {
    param(
        [string]$SourceRoot,
        [string]$VaultRoot
    )

    $sourceFull = [IO.Path]::GetFullPath($SourceRoot).TrimEnd('\') + '\'
    $vaultFull = [IO.Path]::GetFullPath($VaultRoot).TrimEnd('\') + '\'
    $scanErrors = @()
    $files = @(Get-ChildItem -LiteralPath $SourceRoot -File -Recurse -ErrorAction SilentlyContinue -ErrorVariable scanErrors)
    foreach ($scanError in $scanErrors) {
        $script:counts.Errors++
        Write-ExportLog ('ERROR al recorrer OneDrive: ' + $scanError.Exception.Message)
    }
    foreach ($file in $files) {
        $fileFull = [IO.Path]::GetFullPath($file.FullName)
        if ($fileFull.StartsWith($vaultFull, [StringComparison]::OrdinalIgnoreCase)) {
            continue
        }
        try {
            $markdown = Convert-FileToMarkdown -Path $file.FullName
            if ($null -eq $markdown) {
                $script:counts.Skipped++
                Write-ExportLog ('Omitido (formato no compatible): ' + $file.FullName)
                continue
            }
            $relative = $fileFull.Substring($sourceFull.Length)
            $relativeDirectory = [IO.Path]::GetDirectoryName($relative)
            $targetDirectory = Join-Path (Join-Path $VaultRoot 'OneDrive') $relativeDirectory
            if (-not (Test-Path -LiteralPath $targetDirectory)) {
                New-Item -ItemType Directory -Path $targetDirectory -Force | Out-Null
            }
            $targetName = [IO.Path]::GetFileNameWithoutExtension($file.Name) + '.md'
            $targetPath = Get-UniquePath -Path (Join-Path $targetDirectory $targetName)
            $sourceNote = [Uri]::UnescapeDataString($relative.Replace('\', '/'))
            $frontMatter = @(
                '---'
                'tipo: "archivo"'
                ('archivo: ' + (ConvertTo-MarkdownScalar $file.Name))
                ('origen: ' + (ConvertTo-MarkdownScalar $sourceNote))
                ('modificado: ' + (ConvertTo-MarkdownScalar $file.LastWriteTime.ToString('o')))
                '---'
                ''
            ) -join "`r`n"
            [IO.File]::WriteAllText($targetPath, $frontMatter + $markdown, (New-Object Text.UTF8Encoding($false)))
            $script:counts.Files++
            Write-ExportLog ('Convertido: ' + $file.FullName)
        }
        catch {
            $script:counts.Errors++
            Write-ExportLog ('ERROR al convertir ' + $file.FullName + ': ' + $_.Exception.Message)
        }
    }
}

function Export-MailFolder {
    param(
        [object]$Folder,
        [string]$VaultRoot,
        [string]$RelativeFolder
    )

    try {
        $folderItems = $Folder.Items
        for ($index = 1; $index -le $folderItems.Count; $index++) {
            $item = $null
            try {
                $item = $folderItems.Item($index)
                if ($item.Class -ne 43) {
                    continue
                }
                $sentDate = [datetime]$item.ReceivedTime
                $yearFolder = Join-Path $RelativeFolder $sentDate.ToString('yyyy')
                $targetDirectory = Join-Path (Join-Path $VaultRoot 'Outlook') $yearFolder
                if (-not (Test-Path -LiteralPath $targetDirectory)) {
                    New-Item -ItemType Directory -Path $targetDirectory -Force | Out-Null
                }

                $subject = Get-SafeName ([string]$item.Subject)
                $timestamp = $sentDate.ToString('yyyy-MM-dd_HHmmss')
                $targetPath = Get-UniquePath -Path (Join-Path $targetDirectory ($timestamp + ' - ' + $subject + '.md'))
                $attachmentLinks = @()
                if ($item.Attachments.Count -gt 0) {
                    $attachmentDirectory = Join-Path (Join-Path $targetDirectory 'Adjuntos') ([IO.Path]::GetFileNameWithoutExtension($targetPath))
                    New-Item -ItemType Directory -Path $attachmentDirectory -Force | Out-Null
                    for ($attachmentIndex = 1; $attachmentIndex -le $item.Attachments.Count; $attachmentIndex++) {
                        try {
                            $attachment = $item.Attachments.Item($attachmentIndex)
                            $attachmentName = Get-SafeName ([string]$attachment.FileName)
                            $attachmentPath = Get-UniquePath -Path (Join-Path $attachmentDirectory $attachmentName)
                            $attachment.SaveAsFile($attachmentPath)
                            $attachmentLinks += ConvertTo-MarkdownLink -FromFile $targetDirectory -ToFile $attachmentPath -Label $attachmentName
                            $script:counts.Attachments++
                        }
                        catch {
                            $script:counts.Errors++
                            Write-ExportLog ('ERROR al guardar un adjunto de "' + $subject + '": ' + $_.Exception.Message)
                        }
                    }
                }

                $body = [string]$item.Body
                $sender = [string]$item.SenderName
                $to = [string]$item.To
                $cc = [string]$item.CC
                $frontMatter = @(
                    '---'
                    'tipo: "correo"'
                    ('asunto: ' + (ConvertTo-MarkdownScalar $subject))
                    ('de: ' + (ConvertTo-MarkdownScalar $sender))
                    ('para: ' + (ConvertTo-MarkdownScalar $to))
                    ('fecha: ' + (ConvertTo-MarkdownScalar $sentDate.ToString('o')))
                    '---'
                    ''
                    '# ' + $subject
                    ''
                    ('**De:** ' + $sender)
                    ('**Para:** ' + $to)
                )
                if ($cc) {
                    $frontMatter += ('**CC:** ' + $cc)
                }
                $frontMatter += @('', '## Mensaje', '', $body, '')
                if ($attachmentLinks.Count -gt 0) {
                    $frontMatter += @('## Adjuntos', '')
                    foreach ($link in $attachmentLinks) {
                        $frontMatter += ('- ' + $link)
                    }
                }
                [IO.File]::WriteAllText($targetPath, ($frontMatter -join "`r`n"), (New-Object Text.UTF8Encoding($false)))
                $script:counts.Emails++
                Write-ExportLog ('Correo exportado: ' + $subject)
            }
            catch {
                $script:counts.Errors++
                Write-ExportLog ('ERROR al exportar un elemento de correo: ' + $_.Exception.Message)
            }
        }
    }
    catch {
        $script:counts.Errors++
        Write-ExportLog ('ERROR al leer la carpeta de Outlook: ' + $_.Exception.Message)
    }

    try {
        foreach ($childFolder in $Folder.Folders) {
            Export-MailFolder -Folder $childFolder -VaultRoot $VaultRoot -RelativeFolder (Join-Path $RelativeFolder (Get-SafeName ([string]$childFolder.Name)))
        }
    }
    catch {
        $script:counts.Errors++
        Write-ExportLog ('ERROR al recorrer subcarpetas de Outlook: ' + $_.Exception.Message)
    }
}

function Export-OutlookMailbox {
    param([string]$VaultRoot)

    $outlook = New-Object -ComObject Outlook.Application
    $session = $outlook.GetNamespace('MAPI')
    $storeCount = $session.Stores.Count
    for ($storeIndex = 1; $storeIndex -le $storeCount; $storeIndex++) {
        try {
            $store = $session.Stores.Item($storeIndex)
            $root = $store.GetRootFolder()
            $storeName = Get-SafeName ([string]$store.DisplayName)
            Write-ExportLog ('Leyendo buzón: ' + $storeName)
            Export-MailFolder -Folder $root -VaultRoot $VaultRoot -RelativeFolder $storeName
        }
        catch {
            $script:counts.Errors++
            Write-ExportLog ('ERROR al acceder a un buzón de Outlook: ' + $_.Exception.Message)
        }
    }
}

function Close-OfficeApplications {
    if ($script:wordApplication) {
        try { $script:wordApplication.Quit() }
        catch { Write-ExportLog ('AVISO al cerrar Word: ' + $_.Exception.Message) }
        $script:wordApplication = $null
    }
    if ($script:excelApplication) {
        try { $script:excelApplication.Quit() }
        catch { Write-ExportLog ('AVISO al cerrar Excel: ' + $_.Exception.Message) }
        $script:excelApplication = $null
    }
    if ($script:powerPointApplication) {
        try { $script:powerPointApplication.Quit() }
        catch { Write-ExportLog ('AVISO al cerrar PowerPoint: ' + $_.Exception.Message) }
        $script:powerPointApplication = $null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$form = New-Object Windows.Forms.Form
$form.Text = 'Exportar a Obsidian'
$form.Size = New-Object Drawing.Size(760, 580)
$form.MinimumSize = New-Object Drawing.Size(680, 500)
$form.StartPosition = 'CenterScreen'

$sourceLabel = New-Object Windows.Forms.Label
$sourceLabel.Text = 'Carpeta local de OneDrive sincronizada:'
$sourceLabel.Location = New-Object Drawing.Point(16, 18)
$sourceLabel.AutoSize = $true
$form.Controls.Add($sourceLabel)

$sourceBox = New-Object Windows.Forms.TextBox
$sourceBox.Location = New-Object Drawing.Point(18, 42)
$sourceBox.Size = New-Object Drawing.Size(590, 24)
$sourceBox.Anchor = 'Top,Left,Right'
$defaultOneDrive = if ($env:OneDriveCommercial) {
    $env:OneDriveCommercial
}
elseif ($env:OneDriveConsumer) {
    $env:OneDriveConsumer
}
else {
    $env:OneDrive
}
if ($defaultOneDrive) { $sourceBox.Text = $defaultOneDrive }
$form.Controls.Add($sourceBox)

$sourceBrowse = New-Object Windows.Forms.Button
$sourceBrowse.Text = 'Examinar...'
$sourceBrowse.Location = New-Object Drawing.Point(620, 40)
$sourceBrowse.Size = New-Object Drawing.Size(110, 28)
$sourceBrowse.Anchor = 'Top,Right'
$sourceBrowse.Add_Click({
    $dialog = New-Object Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Selecciona la carpeta sincronizada de OneDrive'
    if ($dialog.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) { $sourceBox.Text = $dialog.SelectedPath }
})
$form.Controls.Add($sourceBrowse)

$vaultLabel = New-Object Windows.Forms.Label
$vaultLabel.Text = 'Carpeta destino del vault de Obsidian:'
$vaultLabel.Location = New-Object Drawing.Point(16, 82)
$vaultLabel.AutoSize = $true
$form.Controls.Add($vaultLabel)

$vaultBox = New-Object Windows.Forms.TextBox
$vaultBox.Location = New-Object Drawing.Point(18, 106)
$vaultBox.Size = New-Object Drawing.Size(590, 24)
$vaultBox.Anchor = 'Top,Left,Right'
$form.Controls.Add($vaultBox)

$vaultBrowse = New-Object Windows.Forms.Button
$vaultBrowse.Text = 'Examinar...'
$vaultBrowse.Location = New-Object Drawing.Point(620, 104)
$vaultBrowse.Size = New-Object Drawing.Size(110, 28)
$vaultBrowse.Anchor = 'Top,Right'
$vaultBrowse.Add_Click({
    $dialog = New-Object Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Selecciona el vault de Obsidian'
    if ($dialog.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) { $vaultBox.Text = $dialog.SelectedPath }
})
$form.Controls.Add($vaultBrowse)

$filesCheck = New-Object Windows.Forms.CheckBox
$filesCheck.Text = 'Convertir archivos compatibles de OneDrive'
$filesCheck.Location = New-Object Drawing.Point(18, 150)
$filesCheck.AutoSize = $true
$filesCheck.Checked = $true
$form.Controls.Add($filesCheck)

$mailCheck = New-Object Windows.Forms.CheckBox
$mailCheck.Text = 'Exportar todos los correos de los buzones de Outlook clásico'
$mailCheck.Location = New-Object Drawing.Point(18, 176)
$mailCheck.AutoSize = $true
$mailCheck.Checked = $true
$form.Controls.Add($mailCheck)

$startButton = New-Object Windows.Forms.Button
$startButton.Text = 'Iniciar exportación'
$startButton.Location = New-Object Drawing.Point(18, 210)
$startButton.Size = New-Object Drawing.Size(160, 34)
$form.Controls.Add($startButton)

$logBox = New-Object Windows.Forms.TextBox
$logBox.Location = New-Object Drawing.Point(18, 258)
$logBox.Size = New-Object Drawing.Size(712, 265)
$logBox.Anchor = 'Top,Bottom,Left,Right'
$logBox.Multiline = $true
$logBox.ReadOnly = $true
$logBox.ScrollBars = 'Vertical'
$logBox.Font = New-Object Drawing.Font('Consolas', 9)
$form.Controls.Add($logBox)

$startButton.Add_Click({
    $sourcePath = $sourceBox.Text.Trim()
    $vaultPath = $vaultBox.Text.Trim()
    if (-not $filesCheck.Checked -and -not $mailCheck.Checked) {
        [Windows.Forms.MessageBox]::Show('Selecciona al menos un origen para exportar.', 'Falta una selección', 'OK', 'Warning') | Out-Null
        return
    }
    if ($filesCheck.Checked -and -not (Test-Path -LiteralPath $sourcePath -PathType Container)) {
        [Windows.Forms.MessageBox]::Show('Selecciona una carpeta válida de OneDrive.', 'Ruta no válida', 'OK', 'Warning') | Out-Null
        return
    }
    if ([string]::IsNullOrWhiteSpace($vaultPath)) {
        [Windows.Forms.MessageBox]::Show('Selecciona la carpeta del vault de Obsidian.', 'Falta el destino', 'OK', 'Warning') | Out-Null
        return
    }

    try {
        if (-not (Test-Path -LiteralPath $vaultPath)) {
            New-Item -ItemType Directory -Path $vaultPath -Force | Out-Null
        }
        $vaultPath = [IO.Path]::GetFullPath($vaultPath)
        $script:logPath = Join-Path $vaultPath 'ExportarAObsidian.log'
        $script:counts = @{ Files = 0; Emails = 0; Attachments = 0; Skipped = 0; Errors = 0 }
        $startButton.Enabled = $false
        Write-ExportLog 'Inicio de exportación.'
        if ($filesCheck.Checked) {
            Write-ExportLog 'Leyendo archivos de OneDrive...'
            Export-OneDriveFiles -SourceRoot $sourcePath -VaultRoot $vaultPath
        }
        if ($mailCheck.Checked) {
            Write-ExportLog 'Conectando con Outlook clásico...'
            Export-OutlookMailbox -VaultRoot $vaultPath
        }
        Write-ExportLog ('Finalizado. Archivos: {0}; correos: {1}; adjuntos: {2}; omitidos: {3}; errores: {4}.' -f $script:counts.Files, $script:counts.Emails, $script:counts.Attachments, $script:counts.Skipped, $script:counts.Errors)
        [Windows.Forms.MessageBox]::Show(
            ('Exportación terminada.`r`nArchivos: {0}`r`nCorreos: {1}`r`nAdjuntos: {2}`r`nOmitidos: {3}`r`nErrores: {4}`r`n`r`nConsulta ExportarAObsidian.log para más detalles.' -f $script:counts.Files, $script:counts.Emails, $script:counts.Attachments, $script:counts.Skipped, $script:counts.Errors),
            'Exportación a Obsidian',
            'OK',
            'Information'
        ) | Out-Null
    }
    catch {
        Write-ExportLog ('ERROR: ' + $_.Exception.Message)
        [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Error de exportación', 'OK', 'Error') | Out-Null
    }
    finally {
        Close-OfficeApplications
        $startButton.Enabled = $true
    }
})

[void]$form.ShowDialog()
