Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $PSCommandPath
$pipelineScript = Join-Path $scriptRoot "Convert-OneNoteToMarkdownPipeline.ps1"

if (-not (Test-Path -LiteralPath $pipelineScript -PathType Leaf)) {
    [System.Windows.Forms.MessageBox]::Show(
        "未找到流水线脚本：`r`n$pipelineScript",
        "OneConvert 流水线工具",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}

function Add-LogLine {
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.TextBox]$TextBox,
        [Parameter(Mandatory = $true)][string]$Text
    )

    if ($TextBox.IsDisposed) { return }

    $action = {
        param($tb, $msg)
        $tb.AppendText($msg + [Environment]::NewLine)
    }

    if ($TextBox.InvokeRequired) {
        $null = $TextBox.BeginInvoke($action, @($TextBox, $Text))
    } else {
        & $action $TextBox $Text
    }
}

function Set-ControlEnabled {
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Control[]]$Controls,
        [Parameter(Mandatory = $true)][bool]$Enabled
    )

    foreach ($ctrl in $Controls) {
        if ($null -ne $ctrl -and -not $ctrl.IsDisposed) {
            $ctrl.Enabled = $Enabled
        }
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "OneConvert 流水线工具（.one -> XML -> Markdown）"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(1040, 760)
$form.MinimumSize = New-Object System.Drawing.Size(980, 700)
$form.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)

$layout = New-Object System.Windows.Forms.TableLayoutPanel
$layout.Dock = "Fill"
$layout.ColumnCount = 3
$layout.RowCount = 10
$layout.Padding = New-Object System.Windows.Forms.Padding(12)
$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 150)))
$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 170)))

$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34)))  # intro
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))  # input
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))  # xml out
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))  # md out
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 104))) # options
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 44)))  # buttons
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 26)))  # status
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))  # log
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))  # footer
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 0)))   # spare
$form.Controls.Add($layout)

$lblIntro = New-Object System.Windows.Forms.Label
$lblIntro.Text = "一键处理流程：先将 .one 转换为分层 XML，再转换为分层 Markdown。"
$lblIntro.AutoSize = $true
$lblIntro.Anchor = "Left"
$layout.SetColumnSpan($lblIntro, 3)
$layout.Controls.Add($lblIntro, 0, 0)

$lblInput = New-Object System.Windows.Forms.Label
$lblInput.Text = "输入 .one 文件"
$lblInput.TextAlign = "MiddleLeft"
$lblInput.Dock = "Fill"
$layout.Controls.Add($lblInput, 0, 1)

$txtInput = New-Object System.Windows.Forms.TextBox
$txtInput.Dock = "Fill"
$layout.Controls.Add($txtInput, 1, 1)

$btnBrowseInput = New-Object System.Windows.Forms.Button
$btnBrowseInput.Text = "浏览..."
$btnBrowseInput.Dock = "Fill"
$layout.Controls.Add($btnBrowseInput, 2, 1)

$lblXmlOutput = New-Object System.Windows.Forms.Label
$lblXmlOutput.Text = "XML 输出目录"
$lblXmlOutput.TextAlign = "MiddleLeft"
$lblXmlOutput.Dock = "Fill"
$layout.Controls.Add($lblXmlOutput, 0, 2)

$txtXmlOutput = New-Object System.Windows.Forms.TextBox
$txtXmlOutput.Dock = "Fill"
$layout.Controls.Add($txtXmlOutput, 1, 2)

$xmlOutputButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$xmlOutputButtons.Dock = "Fill"
$xmlOutputButtons.FlowDirection = "LeftToRight"
$xmlOutputButtons.WrapContents = $false

$btnBrowseXmlOutput = New-Object System.Windows.Forms.Button
$btnBrowseXmlOutput.Text = "浏览..."
$btnBrowseXmlOutput.Width = 78
$xmlOutputButtons.Controls.Add($btnBrowseXmlOutput)

$btnOpenXmlOutput = New-Object System.Windows.Forms.Button
$btnOpenXmlOutput.Text = "打开"
$btnOpenXmlOutput.Width = 78
$xmlOutputButtons.Controls.Add($btnOpenXmlOutput)

$layout.Controls.Add($xmlOutputButtons, 2, 2)

$lblMdOutput = New-Object System.Windows.Forms.Label
$lblMdOutput.Text = "Markdown 输出目录"
$lblMdOutput.TextAlign = "MiddleLeft"
$lblMdOutput.Dock = "Fill"
$layout.Controls.Add($lblMdOutput, 0, 3)

$txtMdOutput = New-Object System.Windows.Forms.TextBox
$txtMdOutput.Dock = "Fill"
$layout.Controls.Add($txtMdOutput, 1, 3)

$mdOutputButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$mdOutputButtons.Dock = "Fill"
$mdOutputButtons.FlowDirection = "LeftToRight"
$mdOutputButtons.WrapContents = $false

$btnBrowseMdOutput = New-Object System.Windows.Forms.Button
$btnBrowseMdOutput.Text = "浏览..."
$btnBrowseMdOutput.Width = 78
$mdOutputButtons.Controls.Add($btnBrowseMdOutput)

$btnOpenMdOutput = New-Object System.Windows.Forms.Button
$btnOpenMdOutput.Text = "打开"
$btnOpenMdOutput.Width = 78
$mdOutputButtons.Controls.Add($btnOpenMdOutput)

$layout.Controls.Add($mdOutputButtons, 2, 3)

$lblOptions = New-Object System.Windows.Forms.Label
$lblOptions.Text = "转换选项"
$lblOptions.TextAlign = "MiddleLeft"
$lblOptions.Dock = "Fill"
$layout.Controls.Add($lblOptions, 0, 4)

$optionsBox = New-Object System.Windows.Forms.GroupBox
$optionsBox.Text = "参数"
$optionsBox.Dock = "Fill"
$layout.SetColumnSpan($optionsBox, 2)
$layout.Controls.Add($optionsBox, 1, 4)

$optionsGrid = New-Object System.Windows.Forms.TableLayoutPanel
$optionsGrid.Dock = "Fill"
$optionsGrid.ColumnCount = 6
$optionsGrid.RowCount = 2
$optionsGrid.Padding = New-Object System.Windows.Forms.Padding(8, 10, 8, 6)
$optionsGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$optionsGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$optionsGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$optionsGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$optionsGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$optionsGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$optionsGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34)))
$optionsGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34)))
$optionsBox.Controls.Add($optionsGrid)

$chkIncludeEmpty = New-Object System.Windows.Forms.CheckBox
$chkIncludeEmpty.Text = "包含空页面（XML阶段）"
$chkIncludeEmpty.AutoSize = $true
$chkIncludeEmpty.Margin = New-Object System.Windows.Forms.Padding(0, 6, 16, 0)
$optionsGrid.Controls.Add($chkIncludeEmpty, 0, 0)

$lblTimeout = New-Object System.Windows.Forms.Label
$lblTimeout.Text = "加载超时（秒）"
$lblTimeout.AutoSize = $true
$lblTimeout.Margin = New-Object System.Windows.Forms.Padding(0, 8, 8, 0)
$optionsGrid.Controls.Add($lblTimeout, 1, 0)

$numTimeout = New-Object System.Windows.Forms.NumericUpDown
$numTimeout.Minimum = 5
$numTimeout.Maximum = 300
$numTimeout.Value = 30
$numTimeout.Width = 70
$numTimeout.Margin = New-Object System.Windows.Forms.Padding(0, 4, 16, 0)
$optionsGrid.Controls.Add($numTimeout, 2, 0)

$chkSkipMarkdown = New-Object System.Windows.Forms.CheckBox
$chkSkipMarkdown.Text = "仅生成 XML（跳过 Markdown 阶段）"
$chkSkipMarkdown.AutoSize = $true
$chkSkipMarkdown.Margin = New-Object System.Windows.Forms.Padding(0, 6, 16, 0)
$optionsGrid.Controls.Add($chkSkipMarkdown, 3, 0)
$optionsGrid.SetColumnSpan($chkSkipMarkdown, 3)

$chkCopyAssets = New-Object System.Windows.Forms.CheckBox
$chkCopyAssets.Text = "复制图片资源到 Markdown 输出目录"
$chkCopyAssets.Checked = $true
$chkCopyAssets.AutoSize = $true
$chkCopyAssets.Margin = New-Object System.Windows.Forms.Padding(0, 6, 16, 0)
$optionsGrid.Controls.Add($chkCopyAssets, 0, 1)

$lblImageSyntax = New-Object System.Windows.Forms.Label
$lblImageSyntax.Text = "图片语法"
$lblImageSyntax.AutoSize = $true
$lblImageSyntax.Margin = New-Object System.Windows.Forms.Padding(0, 8, 8, 0)
$optionsGrid.Controls.Add($lblImageSyntax, 1, 1)

$cmbImageSyntax = New-Object System.Windows.Forms.ComboBox
$cmbImageSyntax.DropDownStyle = "DropDownList"
[void]$cmbImageSyntax.Items.Add("markdown")
[void]$cmbImageSyntax.Items.Add("obsidian")
$cmbImageSyntax.SelectedIndex = 0
$cmbImageSyntax.Width = 90
$cmbImageSyntax.Margin = New-Object System.Windows.Forms.Padding(0, 4, 16, 0)
$optionsGrid.Controls.Add($cmbImageSyntax, 2, 1)

$lblAssetDir = New-Object System.Windows.Forms.Label
$lblAssetDir.Text = "资源目录名"
$lblAssetDir.AutoSize = $true
$lblAssetDir.Margin = New-Object System.Windows.Forms.Padding(0, 8, 8, 0)
$optionsGrid.Controls.Add($lblAssetDir, 3, 1)

$txtAssetDir = New-Object System.Windows.Forms.TextBox
$txtAssetDir.Text = "attachment"
$txtAssetDir.Width = 120
$txtAssetDir.Margin = New-Object System.Windows.Forms.Padding(0, 4, 16, 0)
$optionsGrid.Controls.Add($txtAssetDir, 4, 1)

$lblOptionHint = New-Object System.Windows.Forms.Label
$lblOptionHint.Text = "默认图片导出到同级 attachment 文件夹。"
$lblOptionHint.AutoSize = $true
$lblOptionHint.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
$lblOptionHint.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
$optionsGrid.Controls.Add($lblOptionHint, 5, 1)

$buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$buttonPanel.Dock = "Fill"
$buttonPanel.FlowDirection = "LeftToRight"
$buttonPanel.WrapContents = $false
$layout.SetColumnSpan($buttonPanel, 3)
$layout.Controls.Add($buttonPanel, 0, 5)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "开始一键转换"
$btnRun.Width = 140
$btnRun.Height = 32
$buttonPanel.Controls.Add($btnRun)

$btnClearLog = New-Object System.Windows.Forms.Button
$btnClearLog.Text = "清空日志"
$btnClearLog.Width = 120
$btnClearLog.Height = 32
$buttonPanel.Controls.Add($btnClearLog)

$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Text = "退出"
$btnExit.Width = 120
$btnExit.Height = 32
$buttonPanel.Controls.Add($btnExit)

$statusPanel = New-Object System.Windows.Forms.Panel
$statusPanel.Dock = "Fill"
$layout.SetColumnSpan($statusPanel, 3)
$layout.Controls.Add($statusPanel, 0, 6)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "就绪"
$lblStatus.Dock = "Left"
$lblStatus.AutoSize = $true
$statusPanel.Controls.Add($lblStatus)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Dock = "Right"
$progress.Width = 260
$progress.Style = "Continuous"
$statusPanel.Controls.Add($progress)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Dock = "Fill"
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$txtLog.WordWrap = $false
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$layout.SetColumnSpan($txtLog, 3)
$layout.Controls.Add($txtLog, 0, 7)

$lblFooter = New-Object System.Windows.Forms.Label
$lblFooter.Text = "依赖：本机已安装 OneNote（COM）与 Python 3。"
$lblFooter.AutoSize = $true
$lblFooter.Anchor = "Left"
$layout.SetColumnSpan($lblFooter, 3)
$layout.Controls.Add($lblFooter, 0, 8)

$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "OneNote 分区文件 (*.one)|*.one|所有文件 (*.*)|*.*"
$openFileDialog.Title = "选择 OneNote .one 文件"

$folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$folderDialog.Description = "选择目录"
$folderDialog.ShowNewFolderButton = $true

$defaultInput = Join-Path $scriptRoot "新临检.one"
if (Test-Path -LiteralPath $defaultInput) {
    $txtInput.Text = $defaultInput
}
$txtXmlOutput.Text = (Join-Path $scriptRoot "output\\xml")
$txtMdOutput.Text = (Join-Path $scriptRoot "output\\markdown")

$state = [pscustomobject]@{
    IsRunning = $false
    Process = $null
    OutputHandler = $null
    ErrorHandler = $null
    ExitHandler = $null
}

$interactiveControls = @(
    $txtInput, $btnBrowseInput,
    $txtXmlOutput, $btnBrowseXmlOutput, $btnOpenXmlOutput,
    $txtMdOutput, $btnBrowseMdOutput, $btnOpenMdOutput,
    $chkIncludeEmpty, $numTimeout, $chkSkipMarkdown, $chkCopyAssets,
    $cmbImageSyntax, $txtAssetDir, $btnRun
)

function Update-OptionControlState {
    $mdEnabled = -not $chkSkipMarkdown.Checked
    $lblImageSyntax.Enabled = $mdEnabled
    $cmbImageSyntax.Enabled = $mdEnabled
    $chkCopyAssets.Enabled = $mdEnabled
    $btnOpenMdOutput.Enabled = $mdEnabled -and -not $state.IsRunning

    $assetEnabled = $mdEnabled -and $chkCopyAssets.Checked
    $lblAssetDir.Enabled = $assetEnabled
    $txtAssetDir.Enabled = $assetEnabled
}

function Choose-FolderIntoTextBox {
    param([System.Windows.Forms.TextBox]$TargetTextBox)
    if (-not [string]::IsNullOrWhiteSpace($TargetTextBox.Text) -and (Test-Path -LiteralPath $TargetTextBox.Text -PathType Container)) {
        $folderDialog.SelectedPath = $TargetTextBox.Text
    } else {
        $folderDialog.SelectedPath = $scriptRoot
    }
    if ($folderDialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $TargetTextBox.Text = $folderDialog.SelectedPath
    }
}

function Open-DirectoryPath {
    param(
        [string]$PathValue,
        [string]$LabelName
    )

    $path = ($PathValue | ForEach-Object { $_.Trim() })
    if ([string]::IsNullOrWhiteSpace($path)) {
        [System.Windows.Forms.MessageBox]::Show("$LabelName 为空。", "OneConvert 流水线工具") | Out-Null
        return
    }
    if (-not (Test-Path -LiteralPath $path)) {
        [System.Windows.Forms.MessageBox]::Show("$LabelName 尚不存在。`r`n请先执行转换。", "OneConvert 流水线工具") | Out-Null
        return
    }
    Start-Process explorer.exe -ArgumentList "`"$path`""
}

$chkSkipMarkdown.Add_CheckedChanged({ Update-OptionControlState })
$chkCopyAssets.Add_CheckedChanged({ Update-OptionControlState })

$btnBrowseInput.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtInput.Text)) {
        $openFileDialog.InitialDirectory = $scriptRoot
    } else {
        $currentDir = Split-Path -Parent $txtInput.Text
        if (Test-Path -LiteralPath $currentDir -PathType Container) {
            $openFileDialog.InitialDirectory = $currentDir
        }
    }
    if ($openFileDialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtInput.Text = $openFileDialog.FileName
    }
})

$btnBrowseXmlOutput.Add_Click({ Choose-FolderIntoTextBox -TargetTextBox $txtXmlOutput })
$btnBrowseMdOutput.Add_Click({ Choose-FolderIntoTextBox -TargetTextBox $txtMdOutput })
$btnOpenXmlOutput.Add_Click({ Open-DirectoryPath -PathValue $txtXmlOutput.Text -LabelName "XML 输出目录" })
$btnOpenMdOutput.Add_Click({ Open-DirectoryPath -PathValue $txtMdOutput.Text -LabelName "Markdown 输出目录" })

$btnClearLog.Add_Click({
    $txtLog.Clear()
})

$btnExit.Add_Click({
    if ($state.IsRunning -and $null -ne $state.Process -and -not $state.Process.HasExited) {
        $res = [System.Windows.Forms.MessageBox]::Show(
            "当前任务仍在运行，确定要退出吗？",
            "OneConvert 流水线工具",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($res -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }
        try { $state.Process.Kill() } catch {}
    }
    $form.Close()
})

$form.Add_FormClosing({
    if ($state.IsRunning -and $null -ne $state.Process -and -not $state.Process.HasExited) {
        try { $state.Process.Kill() } catch {}
    }
})

$btnRun.Add_Click({
    if ($state.IsRunning) { return }

    $defaultXmlOutputPath = Join-Path $scriptRoot "output\\xml"
    $defaultMdOutputPath = Join-Path $scriptRoot "output\\markdown"

    $inputPath = $txtInput.Text.Trim()
    $xmlOutputPath = $txtXmlOutput.Text.Trim()
    $mdOutputPath = $txtMdOutput.Text.Trim()
    $assetDirName = $txtAssetDir.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($xmlOutputPath)) {
        $xmlOutputPath = $defaultXmlOutputPath
        $txtXmlOutput.Text = $xmlOutputPath
    }
    if (-not $chkSkipMarkdown.Checked -and [string]::IsNullOrWhiteSpace($mdOutputPath)) {
        $mdOutputPath = $defaultMdOutputPath
        $txtMdOutput.Text = $mdOutputPath
    }

    if ([string]::IsNullOrWhiteSpace($inputPath)) {
        [System.Windows.Forms.MessageBox]::Show("请选择 .one 文件。", "OneConvert 流水线工具") | Out-Null
        return
    }
    if (-not (Test-Path -LiteralPath $inputPath -PathType Leaf)) {
        [System.Windows.Forms.MessageBox]::Show("未找到输入文件。`r`n$inputPath", "OneConvert 流水线工具") | Out-Null
        return
    }
    if ([System.IO.Path]::GetExtension($inputPath).ToLowerInvariant() -ne ".one") {
        [System.Windows.Forms.MessageBox]::Show("输入文件必须是 .one 文件。", "OneConvert 流水线工具") | Out-Null
        return
    }
    if (-not $chkSkipMarkdown.Checked -and $chkCopyAssets.Checked -and [string]::IsNullOrWhiteSpace($assetDirName)) {
        [System.Windows.Forms.MessageBox]::Show("启用资源复制时，资源目录名不能为空。", "OneConvert 流水线工具") | Out-Null
        return
    }

    Add-LogLine -TextBox $txtLog -Text ("[{0}] 开始一键转换..." -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))
    Add-LogLine -TextBox $txtLog -Text ("输入文件：{0}" -f $inputPath)
    Add-LogLine -TextBox $txtLog -Text ("XML 输出：{0}" -f $xmlOutputPath)
    if ($chkSkipMarkdown.Checked) {
        Add-LogLine -TextBox $txtLog -Text "Markdown 阶段：已跳过（仅生成 XML）"
    } else {
        Add-LogLine -TextBox $txtLog -Text ("Markdown 输出：{0}" -f $mdOutputPath)
        Add-LogLine -TextBox $txtLog -Text ("图片语法：{0}" -f $cmbImageSyntax.SelectedItem)
        Add-LogLine -TextBox $txtLog -Text ("复制资源：{0}" -f ($(if ($chkCopyAssets.Checked) { "是" } else { "否" })))
        if ($chkCopyAssets.Checked) {
            Add-LogLine -TextBox $txtLog -Text ("资源目录名：{0}" -f $assetDirName)
        }
    }

    $args = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", "`"$pipelineScript`"",
        "-InputOneFile", "`"$inputPath`"",
        "-XmlOutputDirectory", "`"$xmlOutputPath`"",
        "-MarkdownOutputDirectory", "`"$mdOutputPath`"",
        "-LoadTimeoutSeconds", [int]$numTimeout.Value
    )

    if ($chkIncludeEmpty.Checked) { $args += "-IncludeEmptyPages" }
    if ($chkSkipMarkdown.Checked) { $args += "-SkipMarkdownStage" }
    if (-not $chkSkipMarkdown.Checked) {
        $args += @("-ImageSyntax", [string]$cmbImageSyntax.SelectedItem)
        $args += @("-AssetDirectoryName", "`"$assetDirName`"")
        if (-not $chkCopyAssets.Checked) {
            $args += "-CopyAssets:`$false"
        }
    }

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = "powershell.exe"
    $startInfo.Arguments = ($args -join " ")
    $startInfo.WorkingDirectory = $scriptRoot
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $startInfo.StandardErrorEncoding = [System.Text.Encoding]::UTF8

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $startInfo
    $proc.EnableRaisingEvents = $true

    $state.IsRunning = $true
    $state.Process = $proc
    Set-ControlEnabled -Controls $interactiveControls -Enabled $false
    $btnOpenXmlOutput.Enabled = $false
    $btnOpenMdOutput.Enabled = $false
    $lblStatus.Text = "处理中..."
    $progress.Style = "Marquee"
    $progress.MarqueeAnimationSpeed = 25

    $state.OutputHandler = [System.Diagnostics.DataReceivedEventHandler]{
        param($sender, $e)
        if (-not [string]::IsNullOrEmpty($e.Data)) {
            Add-LogLine -TextBox $txtLog -Text $e.Data
        }
    }
    $proc.add_OutputDataReceived($state.OutputHandler)

    $state.ErrorHandler = [System.Diagnostics.DataReceivedEventHandler]{
        param($sender, $e)
        if (-not [string]::IsNullOrEmpty($e.Data)) {
            Add-LogLine -TextBox $txtLog -Text ("[ERR] " + $e.Data)
        }
    }
    $proc.add_ErrorDataReceived($state.ErrorHandler)

    $state.ExitHandler = [System.EventHandler]{
        param($sender, $e)

        $exitCode = $sender.ExitCode
        $completeAction = {
            param($code)
            $state.IsRunning = $false
            Set-ControlEnabled -Controls $interactiveControls -Enabled $true
            $lblStatus.Text = if ($code -eq 0) { "已完成" } else { "失败（退出码=$code）" }
            $progress.Style = "Continuous"
            $progress.MarqueeAnimationSpeed = 0
            $progress.Value = if ($code -eq 0) { 100 } else { 0 }
            Update-OptionControlState
            $btnOpenXmlOutput.Enabled = $true
            if (-not $chkSkipMarkdown.Checked) { $btnOpenMdOutput.Enabled = $true }
            Add-LogLine -TextBox $txtLog -Text ("[{0}] 任务结束，退出码：{1}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $code)
            if ($code -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("处理完成。", "OneConvert 流水线工具") | Out-Null
            }
        }

        if (-not $form.IsDisposed) {
            $null = $form.BeginInvoke($completeAction, @($exitCode))
        }
    }
    $proc.add_Exited($state.ExitHandler)

    try {
        [void]$proc.Start()
        $proc.BeginOutputReadLine()
        $proc.BeginErrorReadLine()
    } catch {
        $state.IsRunning = $false
        Set-ControlEnabled -Controls $interactiveControls -Enabled $true
        Update-OptionControlState
        $lblStatus.Text = "启动失败"
        $progress.Style = "Continuous"
        $progress.MarqueeAnimationSpeed = 0
        $progress.Value = 0
        Add-LogLine -TextBox $txtLog -Text ("[ERR] 启动处理进程失败：{0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show("启动处理进程失败。`r`n$($_.Exception.Message)", "OneConvert 流水线工具") | Out-Null
    }
})

Update-OptionControlState

[void]$form.ShowDialog()
