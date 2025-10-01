<# 
    Merge-PPTX-ByTemplateAndLimits.ps1
    PowerShell 7.5 (Windows). Требуется установленный Microsoft PowerPoint (COM).

    Возможности:
      - GUI-форма (адаптивная, со скроллом) ИЛИ CLI-параметры.
      - Обход большого дерева (десятки тысяч файлов) + быстрый InsertFromFile.
      - Перед каждым блоком слайдов добавляются:
           1) слайд макета "РАЗДЕЛ" (относительный путь или имя папки),
           2) слайд макета "ЗАГОЛОВОК" (имя файла без расширения).
      - Запароленные OOXML → «ЗАПАРОЛЕННЫЕ», битые → «БИТЫЕ».
      - Разделение выходных файлов по лимиту слайдов и/или размера (МБ).
      - Логи (.log + .csv) и сохранение последних путей/настроек в %AppData%\MergePPTX\settings.json.
#>

param(
    [string]$SourceRoot,
    [string]$TemplatePath,
    [string]$TargetRoot,
    [string]$EncryptedRoot,
    [string]$BrokenRoot,
    [ValidateSet('Relative','FolderName')]
    [string]$SectionMode = 'Relative',
    [int]$MaxSlides = 1200,
    [int]$MaxMB = 0,
    [string]$LayoutTitleName,    # ← новое: имя макета «ЗАГОЛОВОК» (если пусто — умолчание)
    [string]$LayoutSectionName,  # ← новое: имя макета «РАЗДЕЛ»   (если пусто — умолчание)
    [switch]$NoGui
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$PP_LAYOUT_TITLE_SLIDE     = 1   # ppLayoutTitle
$PP_LAYOUT_SECTION_HEADER  = 33  # ppLayoutSectionHeader

# Умолчания для имён макетов
$DEFAULT_TITLE_NAME   = 'ЗАГОЛОВОК'
$DEFAULT_SECTION_NAME = 'РАЗДЕЛ'

# --- Перезапуск в STA при необходимости (WinForms/COM) ---
try {
    if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        Write-Host "[WARN] Перезапускаю в STA..."
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = (Get-Command pwsh).Source
        $psi.Arguments = "-NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`" " +
                         ($MyInvocation.BoundParameters.GetEnumerator() | ForEach-Object {
                             "-$($_.Key) `"$($_.Value)`""
                         } | Out-String)
        $psi.UseShellExecute = $true
        [System.Diagnostics.Process]::Start($psi) | Out-Null
        exit
    }
} catch { }

# --- UI + Drawing ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Константы COM PowerPoint ---
$PP_ALERTS_NONE      = 1
$MSO_TRUE            = -1
$MSO_FALSE           = 0
$PP_WINDOW_MINIMIZED = 2

# --- Настройки (кеш путей) ---
$SettingsDir  = Join-Path $env:APPDATA 'MergePPTX'
$SettingsFile = Join-Path $SettingsDir 'settings.json'
[void](New-Item -ItemType Directory -Path $SettingsDir -Force)

function Save-Settings($obj) {
    try { $obj | ConvertTo-Json -Depth 6 | Set-Content -Path $SettingsFile -Encoding UTF8 } catch { }
}
function Load-Settings {
    $defaults = [pscustomobject]@{
        SourceRoot        = $null
        TemplatePath      = $null
        TargetRoot        = $null
        EncryptedRoot     = $null
        BrokenRoot        = $null
        SectionMode       = 'Relative'
        MaxSlides         = 1200
        MaxMB             = 0
        LayoutTitleName   = $DEFAULT_TITLE_NAME    # ← новое поле в настройках
        LayoutSectionName = $DEFAULT_SECTION_NAME  # ← новое поле в настройках
        Remember          = $false
    }
    if (Test-Path $SettingsFile) {
        try {
            $loaded = Get-Content $SettingsFile -Raw | ConvertFrom-Json
            if ($loaded) {
                foreach ($n in $defaults.PSObject.Properties.Name) {
                    if ($loaded.PSObject.Properties.Match($n).Count -eq 0) {
                        Add-Member -InputObject $loaded -MemberType NoteProperty -Name $n -Value $defaults.$n
                    }
                }
                return $loaded
            }
        } catch { }
    }
    return $defaults
}

# --- Логи ---
function New-Logger {
    param([Parameter(Mandatory)] [string] $TargetFolder)

    $logFile = Join-Path $TargetFolder ("MergePPTX_{0:yyyy-MM-dd_HH-mm-ss}.log" -f (Get-Date))
    $csvFile = Join-Path $TargetFolder ("MergePPTX_{0:yyyy-MM-dd_HH-mm-ss}.csv" -f (Get-Date))
    "timestamp,level,message" | Set-Content -Path $csvFile -Encoding UTF8

    $infoSb = {
        param($m)
        "[INFO]  $m" | Tee-Object -FilePath $logFile -Append | Out-Host
    }.GetNewClosure()

    $warnSb = {
        param($m)
        "[WARN]  $m" | Tee-Object -FilePath $logFile -Append | Out-Host
    }.GetNewClosure()

    $errSb = {
        param($m)
        "[ERROR] $m" | Tee-Object -FilePath $logFile -Append | Out-Host
    }.GetNewClosure()

    $rowSb = {
        param($when,$level,$msg)
        $line = ('{0},{1},"{2}"' -f $when.ToString('s'), $level, ($msg -replace '"','""'))
        Add-Content -Path $csvFile -Value $line -Encoding UTF8
    }.GetNewClosure()

    return [pscustomobject]@{
        LogPath = $logFile
        CsvPath = $csvFile
        Info    = $infoSb
        Warn    = $warnSb
        Err     = $errSb
        Row     = $rowSb
    }
}

# --- Утилиты ---
$SupportedExt = @('.pptx','.pptm','.ppsx','.ppsm','.ppt')

function Get-RelativePath {
    param([string]$Root,[string]$FullPath)
    $uriRoot = (Resolve-Path $Root).Path
    $uriFull = (Resolve-Path $FullPath).Path
    $u1 = New-Object System.Uri($uriRoot + [IO.Path]::DirectorySeparatorChar)
    $u2 = New-Object System.Uri($uriFull)
    return [System.Uri]::UnescapeDataString($u1.MakeRelativeUri($u2).ToString()) -replace '/', '\'
}

function Test-IsEncryptedOOXML {
    param([string]$Path)
    try {
        $ext = [IO.Path]::GetExtension($Path).ToLowerInvariant()
        if ($ext -in @('.pptx','.pptm','.ppsx','.ppsm','.docx','.xlsx','.xlsm','.docm')) {
            $fs = [IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
            try {
                $buf = New-Object byte[] 8
                $null = $fs.Read($buf,0,8)
                $sig = ($buf | ForEach-Object { $_.ToString('X2') }) -join ''
                if ($sig.StartsWith('D0CF11E0A1B11AE1')) { return $true } else { return $false } # CFBF → encrypted OOXML
            } finally { $fs.Close() }
        }
        return $false
    } catch { return $false }
}

function Move-PreservingRelative {
    param(
        [Parameter(Mandatory)] [string] $File,
        [Parameter(Mandatory)] [string] $DestRoot,
        [Parameter(Mandatory)] [string] $SourceRoot
    )
    try {
        $rel = Get-RelativePath -Root $SourceRoot -FullPath $File
        $dest = Join-Path $DestRoot $rel
        $destDir = Split-Path $dest -Parent
        [void](New-Item -ItemType Directory -Path $destDir -Force)
        Move-Item -LiteralPath $File -Destination $dest -Force
        return $dest
    } catch {
        $name = [IO.Path]::GetFileName($File)
        $fallback = Join-Path $DestRoot $name
        Move-Item -LiteralPath $File -Destination $fallback -Force
        return $fallback
    }
}

function Normalize-Name([string]$s) {
    if ([string]::IsNullOrEmpty($s)) { return "" }
    $s = $s -replace '[\u00A0\u202F]', ' '
    $s = $s -replace '[\u200B\u200C\u200D\uFEFF\u00AD]', ''
    $s = $s.Trim()
    $s = $s -replace '\s+', ' '
    return $s.ToLowerInvariant()
}

function Get-TemplateMasters($Presentation) {
    $masters = @()
    try {
        if ($Presentation.Designs) {
            foreach ($d in $Presentation.Designs) {
                try { if ($d.SlideMaster) { $masters += $d.SlideMaster } } catch {}
            }
        }
    } catch {}
    try { if ($Presentation.SlideMaster) { $masters += $Presentation.SlideMaster } } catch {}
    return $masters | Where-Object { $_ -ne $null } | Select-Object -Unique
}

function Dump-TemplateLayouts($Presentation, $logger=$null) {
    $i = 0
    foreach ($m in (Get-TemplateMasters -Presentation $Presentation)) {
        $i++
        $msg = "Считываю шаблоны из $($Presentation.Name)"
        if ($logger) { $logger.Invoke($msg) } else { Write-Host $msg }
        $msg = "Master #${i}: Макеты:"
        if ($logger) { $logger.Invoke($msg) } else { Write-Host $msg }
        try {
            $idx = 0
            foreach ($cl in $m.CustomLayouts) {
                $idx++
                $name = ($cl.Name -as [string])
                $line = ("  [{0,2}] {1}" -f $idx, $name)
                if ($logger) { $logger.Invoke($line) } else { Write-Host $line }
            }
        } catch {}
    }
}

function Get-CustomLayoutByName {
    param(
        $Presentation,
        [Parameter(Mandatory)][string]$Name,
        [string[]]$Also = @()
    )
    $targets = @($Name) + $Also
    $targetsNorm = $targets | ForEach-Object { Normalize-Name $_ } | Where-Object { $_ }

    foreach ($m in (Get-TemplateMasters -Presentation $Presentation)) {
        foreach ($cl in $m.CustomLayouts) {
            $nm = Normalize-Name ($cl.Name -as [string])
            if ($targetsNorm -contains $nm) { return $cl }
        }
    }
    foreach ($m in (Get-TemplateMasters -Presentation $Presentation)) {
        foreach ($cl in $m.CustomLayouts) {
            $nm = Normalize-Name ($cl.Name -as [string])
            foreach ($t in $targetsNorm) {
                if ($nm.Contains($t) -or $nm.StartsWith($t)) { return $cl }
            }
        }
    }
    return $null
}

function Set-SlideTitle {
    param($Slide,[string]$Text)
    try {
        if ($Slide.Shapes.Title) {
            $Slide.Shapes.Title.TextFrame.TextRange.Text = $Text
            return
        }
    } catch { }
    foreach ($sh in $Slide.Shapes) {
        try {
            if ($sh.HasTextFrame -ne $null -and $sh.HasTextFrame -ne 0) {
                if ($sh.TextFrame.HasText -ne $null) {
                    $sh.TextFrame.TextRange.Text = $Text
                    return
                }
            }
        } catch { }
    }
}

function New-OutputFromTemplate {
    param($PPApp,[string]$TemplatePath,[string]$OutPath)
    $pres = $PPApp.Presentations.Open($TemplatePath, $false, $true, $false)
    $pres.SaveAs($OutPath)
    return $pres
}

function Add-LeadSlides {
    param(
        $DestPres, $LayoutSection, $LayoutTitle,
        [string]$SectionTitleText, [string]$FileTitleText
    )
    $idx = $DestPres.Slides.Count + 1
    if ($LayoutSection) {
        $s1 = $DestPres.Slides.AddSlide($idx, $LayoutSection)
    } else {
        $s1 = $DestPres.Slides.Add($idx, $PP_LAYOUT_SECTION_HEADER)
    }
    Set-SlideTitle -Slide $s1 -Text $SectionTitleText

    $idx2 = $DestPres.Slides.Count + 1
    if ($LayoutTitle) {
        $s2 = $DestPres.Slides.AddSlide($idx2, $LayoutTitle)
    } else {
        $s2 = $DestPres.Slides.Add($idx2, $PP_LAYOUT_TITLE_SLIDE)
    }
    Set-SlideTitle -Slide $s2 -Text $FileTitleText
}

function Insert-AllSlidesFromFile {
    param(
        $DestPres,
        [Parameter(Mandatory)][string]$SourceFile
    )
    $insertAfter = $DestPres.Slides.Count
    $src = $null
    try {
        $src = $script:pp.Presentations.Open($SourceFile, $true, $true, $false)
        $n = [int]$src.Slides.Count
    } finally {
        if ($src) { try { $src.Close() } catch {} }
    }
    if ($n -lt 1) { return 0 }
    $null = $DestPres.Slides.InsertFromFile($SourceFile, $insertAfter, 1, $n)
    return ($DestPres.Slides.Count - $insertAfter)
}

# --- GUI ---
function Show-InputForm {
    $settings = Load-Settings

    $form = New-Object System.Windows.Forms.Form
    $form.Size = New-Object System.Drawing.Size(980, 500)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.MaximizeBox = $false
    $form.SizeGripStyle = [System.Windows.Forms.SizeGripStyle]::Hide
    $form.MinimumSize = $form.Size
    $form.MaximumSize = $form.Size
    $form.AutoScroll = $true
    $form.Text = "Объединение презентаций по шаблону"
    $form.StartPosition = "CenterScreen"
    $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
    $form.TopMost = $false

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'
    $layout.AutoScroll = $true
    $layout.ColumnCount = 3
    $layout.RowCount = 0
    $layout.Padding = '10,10,10,10'
    $layout.GrowStyle = [System.Windows.Forms.TableLayoutPanelGrowStyle]::AddRows
    [void]$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 270)))
    [void]$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 40)))

    function New-Label([string]$text) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $text; $l.AutoSize = $true; $l.Margin = '0,6,8,6'
        return $l
    }
    function New-PathRow([string]$caption,[ref]$tbOut,[switch]$File,[string]$filter="PowerPoint (*.pptx;*.potx)|*.pptx;*.potx|All files (*.*)|*.*") {
        $lbl = New-Label $caption
        $tb  = New-Object System.Windows.Forms.TextBox; $tb.Anchor='Left,Right'; $tb.Margin='0,3,6,3'; $tb.Width=600
        $btn = New-Object System.Windows.Forms.Button;  $btn.Text='...'; $btn.Width=32; $btn.Margin='0,0,0,0'; $btn.Anchor='Right'
        if ($File) { $ofd = New-Object System.Windows.Forms.OpenFileDialog; $ofd.Filter=$filter; $btn.Tag=@{TB=$tb;Mode='file';OFD=$ofd} }
        else { $fbd = New-Object System.Windows.Forms.FolderBrowserDialog; $btn.Tag=@{TB=$tb;Mode='folder';FBD=$fbd} }
        [void]$btn.Add_Click({
            $tag = $this.Tag; $tb2 = [System.Windows.Forms.TextBox]$tag.TB
            if ($tag.Mode -eq 'file') { $ofd = $tag.OFD; if ($tb2.Text) { $ofd.FileName = $tb2.Text }; if ($ofd.ShowDialog() -eq 'OK') { $tb2.Text = $ofd.FileName } }
            else { $fbd = $tag.FBD; if ($tb2.Text -and (Test-Path $tb2.Text)) { $fbd.SelectedPath = $tb2.Text }; if ($fbd.ShowDialog() -eq 'OK') { $tb2.Text = $fbd.SelectedPath } }
        })
        $row = $layout.RowCount
        [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
        [void]$layout.Controls.Add($lbl,0,$row); [void]$layout.Controls.Add($tb,1,$row); [void]$layout.Controls.Add($btn,2,$row)
        $layout.RowCount++
        $tbOut.Value = $tb
    }
    function New-TextRow([string]$caption,[ref]$tbOut,[string]$placeholder="") {
        $lbl = New-Label $caption
        $tb  = New-Object System.Windows.Forms.TextBox
        $tb.Anchor='Left,Right'; $tb.Margin='0,3,6,3'; $tb.Width=600
        if ($placeholder) { $tb.Text = $placeholder }
        $row = $layout.RowCount
        [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
        [void]$layout.Controls.Add($lbl,0,$row); [void]$layout.Controls.Add($tb,1,$row)
        $layout.RowCount++
        $tbOut.Value = $tb
    }

    $tbSrc=$null; New-PathRow "1) Директория-источник:" ([ref]$tbSrc)
    $tbTpl=$null; New-PathRow "2) Презентация-шаблон (.pptx/.potx):" ([ref]$tbTpl) -File
    $tbDst=$null; New-PathRow "3) Целевая папка результатов:" ([ref]$tbDst)
    $tbEnc=$null; New-PathRow "4) Папка «ЗАПАРОЛЕННЫЕ»:" ([ref]$tbEnc)
    $tbBad=$null; New-PathRow "5) Папка «БИТЫЕ»:" ([ref]$tbBad)

    $lblMode = New-Label "6) Заголовок слайда разделителя:"
    $modePanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $modePanel.Dock='Fill'; $modePanel.AutoSize=$true; $modePanel.WrapContents=$false; $modePanel.Margin='0,3,6,3'
    $rbRel  = New-Object System.Windows.Forms.RadioButton; $rbRel.Text="Относительный путь от директории-источника"; $rbRel.AutoSize=$true; $rbRel.Margin='0,3,25,3'
    $rbName = New-Object System.Windows.Forms.RadioButton; $rbName.Text="Только имя родительской папки";   $rbName.AutoSize=$true; $rbName.Margin='0,3,0,3'
    [void]$modePanel.Controls.Add($rbRel); [void]$modePanel.Controls.Add($rbName)
    $row = $layout.RowCount
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$layout.Controls.Add($lblMode,0,$row); [void]$layout.Controls.Add($modePanel,1,$row); $layout.RowCount++

    $lblLimits = New-Label "Лимиты на выходной файл:"
    $limitsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $limitsPanel.Dock='Fill'; $limitsPanel.AutoSize=$true; $limitsPanel.WrapContents=$false; $limitsPanel.Margin='0,3,6,3'
    $lblMaxS = New-Label "Макс. слайдов:"; $numMaxS = New-Object System.Windows.Forms.NumericUpDown
    $numMaxS.Minimum=1; $numMaxS.Maximum=100000; $numMaxS.Value=1200; $numMaxS.Width=100; $numMaxS.Margin='6,0,25,0'
    $lblMaxMB = New-Label "Макс. размер (МБ, 0 = нет лимита):"; $numMaxMB = New-Object System.Windows.Forms.NumericUpDown
    $numMaxMB.Minimum=0; $numMaxMB.Maximum=10000; $numMaxMB.Value=0; $numMaxMB.Width=100; $numMaxMB.Margin='6,0,0,0'
    [void]$limitsPanel.Controls.Add($lblMaxS); [void]$limitsPanel.Controls.Add($numMaxS); [void]$limitsPanel.Controls.Add($lblMaxMB); [void]$limitsPanel.Controls.Add($numMaxMB)
    $row = $layout.RowCount
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$layout.Controls.Add($lblLimits,0,$row); [void]$layout.Controls.Add($limitsPanel,1,$row); $layout.RowCount++

    # --- Новые поля для имён макетов ---
    $tbTitleName=$null
    New-TextRow "7) Название макета «ЗАГОЛОВОК» (по умолчанию: $DEFAULT_TITLE_NAME):" ([ref]$tbTitleName) ($settings.LayoutTitleName)

    $tbSectionName=$null
    New-TextRow "8) Название макета «РАЗДЕЛ» (по умолчанию: $DEFAULT_SECTION_NAME):" ([ref]$tbSectionName) ($settings.LayoutSectionName)

    $lblRemember = New-Label ""
    $chkRemember = New-Object System.Windows.Forms.CheckBox; $chkRemember.Text="Запомнить эти параметры на будущее"; $chkRemember.AutoSize=$true; $chkRemember.Margin='0,6,6,6'
    $row = $layout.RowCount
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$layout.Controls.Add($lblRemember,0,$row); [void]$layout.Controls.Add($chkRemember,1,$row); $layout.RowCount++

    $btnPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $btnPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
    $btnPanel.WrapContents  = $false
    $btnPanel.AutoSize = $false
    $btnPanel.Height   = 40
    $btnPanel.Width    = 250
    $btnPanel.Margin   = '0,12,6,6'
    $btnPanel.Anchor   = 'Right'

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = 'Старт'
    $btnRun.Size = New-Object System.Drawing.Size(120,36)
    $btnRun.Margin = '6,0,0,0'

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Отмена'
    $btnCancel.Size = New-Object System.Drawing.Size(110,36)
    $btnCancel.Margin = '0,0,6,0'

    [void]$btnPanel.Controls.Add($btnRun)
    [void]$btnPanel.Controls.Add($btnCancel)

    $row = $layout.RowCount
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$layout.Controls.Add((New-Object System.Windows.Forms.Label),0,$row)
    [void]$layout.Controls.Add($btnPanel,1,$row)
    $layout.RowCount++

    [void]$form.Controls.Add($layout)

    # Подстановка сохранённых
    if ($settings.SourceRoot)    { $tbSrc.Text = $settings.SourceRoot }
    if ($settings.TemplatePath)  { $tbTpl.Text = $settings.TemplatePath }
    if ($settings.TargetRoot)    { $tbDst.Text = $settings.TargetRoot }
    if ($settings.EncryptedRoot) { $tbEnc.Text = $settings.EncryptedRoot }
    if ($settings.BrokenRoot)    { $tbBad.Text = $settings.BrokenRoot }
    if ($settings.SectionMode -eq 'Relative') { $rbRel.Checked = $true } elseif ($settings.SectionMode -eq 'FolderName') { $rbName.Checked = $true } else { $rbRel.Checked = $true }
    if ($settings.MaxSlides) { $numMaxS.Value = [decimal]$settings.MaxSlides }
    if ($settings.MaxMB -ne $null) { $numMaxMB.Value = [decimal]$settings.MaxMB }
    if ($settings.Remember) { $chkRemember.Checked = $true }

    $resultData = @{
        DialogResult      = $null
        SourceRoot        = $null
        TemplatePath      = $null
        TargetRoot        = $null
        EncryptedRoot     = $null
        BrokenRoot        = $null
        SectionMode       = 'Relative'
        MaxSlides         = 1200
        MaxMB             = 0
        LayoutTitleName   = $DEFAULT_TITLE_NAME
        LayoutSectionName = $DEFAULT_SECTION_NAME
        Remember          = $false
    }

    [void]$btnCancel.Add_Click({ $resultData.DialogResult = 'Cancel'; $form.Close() })
    [void]$btnRun.Add_Click({
        if (-not (Test-Path $tbSrc.Text)) { [void][System.Windows.Forms.MessageBox]::Show("Укажите корректную директорию-источник"); return }
        if (-not (Test-Path $tbTpl.Text)) { [void][System.Windows.Forms.MessageBox]::Show("Укажите корректный путь к шаблону"); return }
        if (-not (Test-Path $tbDst.Text)) { [void][System.Windows.Forms.MessageBox]::Show("Укажите корректную целевую папку результатов"); return }
        if (-not (Test-Path $tbEnc.Text)) { [void][System.Windows.Forms.MessageBox]::Show("Укажите корректную папку «ЗАПАРОЛЕННЫЕ»"); return }
        if (-not (Test-Path $tbBad.Text)) { [void][System.Windows.Forms.MessageBox]::Show("Укажите корректную папку «БИТЫЕ»"); return }
        if (-not ($rbRel.Checked -or $rbName.Checked)) { [void][System.Windows.Forms.MessageBox]::Show("Выберите режим заголовка слайда-разделителя"); return }

        $resultData.SourceRoot        = $tbSrc.Text
        $resultData.TemplatePath      = $tbTpl.Text
        $resultData.TargetRoot        = $tbDst.Text
        $resultData.EncryptedRoot     = $tbEnc.Text
        $resultData.BrokenRoot        = $tbBad.Text
        $resultData.SectionMode       = if ($rbRel.Checked) { 'Relative' } else { 'FolderName' }
        $resultData.MaxSlides         = [int]$numMaxS.Value
        $resultData.MaxMB             = [int]$numMaxMB.Value
        $resultData.LayoutTitleName   = if ([string]::IsNullOrWhiteSpace($tbTitleName.Text))   { $DEFAULT_TITLE_NAME }   else { $tbTitleName.Text }
        $resultData.LayoutSectionName = if ([string]::IsNullOrWhiteSpace($tbSectionName.Text)) { $DEFAULT_SECTION_NAME } else { $tbSectionName.Text }
        $resultData.Remember          = $chkRemember.Checked
        $resultData.DialogResult      = 'OK'

        if ($resultData.Remember) {
            Save-Settings([pscustomobject]@{
                SourceRoot        = $resultData.SourceRoot
                TemplatePath      = $resultData.TemplatePath
                TargetRoot        = $resultData.TargetRoot
                EncryptedRoot     = $resultData.EncryptedRoot
                BrokenRoot        = $resultData.BrokenRoot
                SectionMode       = $resultData.SectionMode
                MaxSlides         = $resultData.MaxSlides
                MaxMB             = $resultData.MaxMB
                LayoutTitleName   = $resultData.LayoutTitleName
                LayoutSectionName = $resultData.LayoutSectionName
                Remember          = $true
            })
        }
        $form.Close()
    })

    [void]$form.ShowDialog()
    if ($resultData.DialogResult -ne 'OK') { return $null }
    return [pscustomobject]$resultData
}

# --------------------------- СБОР ВХОДНЫХ ДАННЫХ ---------------------------

if (-not $NoGui -and (
    [string]::IsNullOrWhiteSpace($SourceRoot) -or
    [string]::IsNullOrWhiteSpace($TemplatePath) -or
    [string]::IsNullOrWhiteSpace($TargetRoot) -or
    [string]::IsNullOrWhiteSpace($EncryptedRoot) -or
    [string]::IsNullOrWhiteSpace($BrokenRoot)
)) {
    $inputData = Show-InputForm
    if ($null -eq $inputData) { Write-Host "Отменено пользователем."; return }
    $SourceRoot        = $inputData.SourceRoot
    $TemplatePath      = $inputData.TemplatePath
    $TargetRoot        = $inputData.TargetRoot
    $EncryptedRoot     = $inputData.EncryptedRoot
    $BrokenRoot        = $inputData.BrokenRoot
    $SectionMode       = $inputData.SectionMode
    $MaxSlides         = [int]$inputData.MaxSlides
    $MaxMB             = [int]$inputData.MaxMB
    $LayoutTitleName   = $inputData.LayoutTitleName
    $LayoutSectionName = $inputData.LayoutSectionName
} else {
    # Если GUI не запускали: подставим умолчания, если имена не заданы параметрами
    if ([string]::IsNullOrWhiteSpace($LayoutTitleName))   { $LayoutTitleName   = $DEFAULT_TITLE_NAME }
    if ([string]::IsNullOrWhiteSpace($LayoutSectionName)) { $LayoutSectionName = $DEFAULT_SECTION_NAME }
}

# Жёсткая проверка
$required = @(
    @{Name='SourceRoot';    Val=$SourceRoot },
    @{Name='TemplatePath';  Val=$TemplatePath },
    @{Name='TargetRoot';    Val=$TargetRoot },
    @{Name='EncryptedRoot'; Val=$EncryptedRoot },
    @{Name='BrokenRoot';    Val=$BrokenRoot }
)
foreach ($r in $required) {
    if ([string]::IsNullOrWhiteSpace($r.Val)) { throw "Поле '$($r.Name)' пустое. Заполните его (GUI или параметрами)." }
}

# Создание папок
[void](New-Item -ItemType Directory -Path $TargetRoot -Force)
[void](New-Item -ItemType Directory -Path $EncryptedRoot -Force)
[void](New-Item -ItemType Directory -Path $BrokenRoot -Force)

# --------------------------- ОСНОВНОЙ ПРОЦЕСС ---------------------------

$log = New-Logger -TargetFolder $TargetRoot
$log.Info.Invoke("=== MERGE START ===")
$log.Info.Invoke("SourceRoot: $SourceRoot")
$log.Info.Invoke("Template:   $TemplatePath")
$log.Info.Invoke("TargetRoot: $TargetRoot")
$log.Info.Invoke("Encrypted:  $EncryptedRoot")
$log.Info.Invoke("Broken:     $BrokenRoot")
$log.Info.Invoke("SectionMode: $SectionMode; MaxSlides: $MaxSlides; MaxMB: $MaxMB")
$log.Info.Invoke("Layout names → Title: '$LayoutTitleName'; Section: '$LayoutSectionName'")

$stopwatch   = [System.Diagnostics.Stopwatch]::StartNew()
$processed   = 0
$movedEnc    = 0
$movedBad    = 0
$addedSlides = 0
$outIndex    = 1

# PowerPoint COM
$pp = $null
$dest = $null
$layoutSection = $null
$layoutTitle   = $null

try {
    $pp = New-Object -ComObject PowerPoint.Application
    $pp.WindowState   = $PP_WINDOW_MINIMIZED
    $pp.DisplayAlerts = $PP_ALERTS_NONE
} catch {
    $log.Err.Invoke("Не удалось запустить PowerPoint (COM): $($_.Exception.Message)")
    throw
}

function Open-NewOutput {
    param([int]$Index)
    $outName = "MERGED_{0:yyyyMMdd_HHmmss}_{1:D3}.pptx" -f (Get-Date), $Index
    $outPath = Join-Path $TargetRoot $outName
    $log.Info.Invoke("Открываю новый выходной файл: $outName")
    $pres = New-OutputFromTemplate -PPApp $pp -TemplatePath $TemplatePath -OutPath $outPath

    Dump-TemplateLayouts -Presentation $pres -logger $log.Info

    # Используем заданные пользователем имена (или умолчания), + синонимы
    $wantSection = if ([string]::IsNullOrWhiteSpace($LayoutSectionName)) { $DEFAULT_SECTION_NAME } else { $LayoutSectionName }
    $wantTitle   = if ([string]::IsNullOrWhiteSpace($LayoutTitleName))   { $DEFAULT_TITLE_NAME }   else { $LayoutTitleName }

    $script:layoutSection = Get-CustomLayoutByName -Presentation $pres -Name $wantSection -Also @('Заголовок 2 уровня','Section Header','Раздел')
    $script:layoutTitle   = Get-CustomLayoutByName -Presentation $pres -Name $wantTitle   -Also @('Заголовок 1 уровня','Title Slide','Title','Титул')

    if (-not $script:layoutSection -or -not $script:layoutTitle) {
        $log.Warn.Invoke("Точный макет '$wantSection' и/или '$wantTitle' не найден: используем резервные способы.")
    }
    return $pres
}

function Maybe-RotateOutput {
    param($Pres)
    $needRotate = $false
    if ($Pres.Slides.Count -ge $MaxSlides) { $needRotate = $true }
    if (-not $needRotate -and $MaxMB -gt 0) {
        try {
            $path = $Pres.FullName
            $Pres.Save()
            $sizeMB = [math]::Round((Get-Item $path).Length / 1MB, 2)
            if ($sizeMB -ge $MaxMB) { $needRotate = $true }
        } catch { }
    }
    if ($needRotate) {
        $global:outIndex++
        $log.Info.Invoke("Достигнут лимит. Сохраняю и начинаю новый файл...")
        $Pres.Save(); $Pres.Close()
        $script:dest = Open-NewOutput -Index $global:outIndex
    }
}

$dest = Open-NewOutput -Index $outIndex

$filesEnum = Get-ChildItem -LiteralPath $SourceRoot -Recurse -File -ErrorAction SilentlyContinue |
    Where-Object { $SupportedExt -contains $_.Extension.ToLowerInvariant() }

foreach ($f in $filesEnum) {
    $processed++
    $full = $f.FullName

    $isEncrypted = $false
    try { $isEncrypted = Test-IsEncryptedOOXML -Path $full } catch { $isEncrypted = $false }
    if ($isEncrypted) {
        $moved = Move-PreservingRelative -File $full -DestRoot $EncryptedRoot -SourceRoot $SourceRoot
        $movedEnc++
        $log.Row.Invoke((Get-Date),'encrypted', "Moved to $moved")
        Write-Progress -Activity "Обработка презентаций" -Status "Пароль: $($f.Name) → ЗАПАРОЛЕННЫЕ" -PercentComplete 0
        continue
    }

    $fileTitle   = [IO.Path]::GetFileNameWithoutExtension($full)
    $sectionTitle = switch ($SectionMode) {
        'FolderName' { Split-Path $full -Parent | Split-Path -Leaf }
        default      { Get-RelativePath -Root $SourceRoot -FullPath (Split-Path $full -Parent) }
    }

    try {
        Add-LeadSlides -DestPres $dest -LayoutSection $layoutSection -LayoutTitle $layoutTitle `
                       -SectionTitleText $sectionTitle -FileTitleText $fileTitle

        $added = Insert-AllSlidesFromFile -DestPres $dest -SourceFile $full
        $addedSlides += $added
        $log.Row.Invoke((Get-Date),'ok', "Added $added slides from $full; section='$sectionTitle'; title='$fileTitle'")

        Maybe-RotateOutput -Pres $dest
    } catch {
        try {
            $moved = Move-PreservingRelative -File $full -DestRoot $BrokenRoot -SourceRoot $SourceRoot
            $movedBad++
            $log.Row.Invoke((Get-Date),'broken', "Moved to $moved; error=$($_.Exception.Message)")
        } catch {
            $log.Err.Invoke("Ошибка и при переносе в «БИТЫЕ»: $full :: $($_.Exception.Message)")
        }
        continue
    }

    if ($processed % 50 -eq 0) {
        $elapsed = $stopwatch.Elapsed
        $rate = if ($elapsed.TotalSeconds -gt 0) { [math]::Round($processed / $elapsed.TotalSeconds, 2) } else { 0 }
        Write-Progress -Activity "Файлы: $processed; добавлено слайдов: $addedSlides" `
                       -Status ("Скорость ~{0} файлов/с" -f $rate) -PercentComplete 0
    }
}

try { if ($dest) { $dest.Save(); $dest.Close() } } catch { $log.Warn.Invoke("Не удалось нормально закрыть текущую презентацию: $($_.Exception.Message)") }
try { if ($pp) { $pp.DisplayAlerts = $PP_ALERTS_NONE; $pp.Quit() } } catch { }

[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

$stopwatch.Stop()
$log.Info.Invoke("=== MERGE DONE ===")
$log.Info.Invoke("Обработано файлов: $processed; добавлено слайдов: $addedSlides; запароленных: $movedEnc; битых: $movedBad; Время: $($stopwatch.Elapsed)")
Write-Host "`nГотово.`nЛоги: $($log.LogPath)`nCSV-отчёт: $($log.CsvPath)"
