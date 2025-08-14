Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Add-Log {
    param([string]$msg)
    $tbLog.AppendText( (Get-Date -Format 'HH:mm:ss') + " - $msg`r`n" )
    $tbLog.SelectionStart = $tbLog.Text.Length
    $tbLog.ScrollToCaret()
}

if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    $newProcess = New-Object System.Diagnostics.ProcessStartInfo "powershell";
    $newProcess.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$($myInvocation.MyCommand.Definition)`"";
    $newProcess.Verb = "runas";
    [System.Diagnostics.Process]::Start($newProcess);
    exit;
}

# --- GUI DESIGN ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "1C 8.3 Agent/Remote Server Manipulation"
$form.Size = New-Object System.Drawing.Size(560,540)
$form.StartPosition = "CenterScreen"

# Platform Path
$lblPlatform = New-Object System.Windows.Forms.Label
$lblPlatform.Text = "Папка с платформой:"
$lblPlatform.Location = New-Object System.Drawing.Point(10,20)
$lblPlatform.Size = New-Object System.Drawing.Size(130,22)
$form.Controls.Add($lblPlatform)

$txtPlatform = New-Object System.Windows.Forms.TextBox
$txtPlatform.Location = New-Object System.Drawing.Point(145,18)
$txtPlatform.Width = 300
$form.Controls.Add($txtPlatform)

$btnBrowsePlatform = New-Object System.Windows.Forms.Button
$btnBrowsePlatform.Text = "..."
$btnBrowsePlatform.Location = New-Object System.Drawing.Point(455,17)
$btnBrowsePlatform.Width = 30
$form.Controls.Add($btnBrowsePlatform)

# Base Folder (upper-level folder for srvinfo creation)
$lblBaseFolder = New-Object System.Windows.Forms.Label
$lblBaseFolder.Text = "Верхнеуровневая папка:"
$lblBaseFolder.Location = New-Object System.Drawing.Point(10,55)
$lblBaseFolder.Size = New-Object System.Drawing.Size(130,22)
$form.Controls.Add($lblBaseFolder)

$txtBaseFolder = New-Object System.Windows.Forms.TextBox
$txtBaseFolder.Location = New-Object System.Drawing.Point(145,53)
$txtBaseFolder.Width = 300
$form.Controls.Add($txtBaseFolder)

$btnBrowseBaseFolder = New-Object System.Windows.Forms.Button
$btnBrowseBaseFolder.Text = "..."
$btnBrowseBaseFolder.Location = New-Object System.Drawing.Point(455,52)
$btnBrowseBaseFolder.Width = 30
$form.Controls.Add($btnBrowseBaseFolder)

# StartPort
$lblPort = New-Object System.Windows.Forms.Label
$lblPort.Text = "Стартовый порт:"
$lblPort.Location = New-Object System.Drawing.Point(10,90)
$form.Controls.Add($lblPort)

$numPort = New-Object System.Windows.Forms.NumericUpDown
$numPort.Location = New-Object System.Drawing.Point(145,88)
$numPort.Minimum = 1
$numPort.Maximum = 65000
$numPort.Value = 29
$form.Controls.Add($numPort)

# DisplayName additions for Ragent
$lblRagentSuffix = New-Object System.Windows.Forms.Label
$lblRagentSuffix.Text = "Приписка к DisplayName ragent:"
$lblRagentSuffix.Location = New-Object System.Drawing.Point(10,125)
$lblRagentSuffix.Size = New-Object System.Drawing.Size(180,22)
$form.Controls.Add($lblRagentSuffix)

$txtRagentSuffix = New-Object System.Windows.Forms.TextBox
$txtRagentSuffix.Location = New-Object System.Drawing.Point(195,123)
$txtRagentSuffix.Width = 290
$form.Controls.Add($txtRagentSuffix)

# DisplayName additions for Ras (Remote Server)
$lblRasSuffix = New-Object System.Windows.Forms.Label
$lblRasSuffix.Text = "Приписка к DisplayName ras:"
$lblRasSuffix.Location = New-Object System.Drawing.Point(10,160)
$lblRasSuffix.Size = New-Object System.Drawing.Size(180,22)
$form.Controls.Add($lblRasSuffix)

# Debug флажок
$chkDebug = New-Object System.Windows.Forms.CheckBox
$chkDebug.Text = "Включить debug (-debug)"
$chkDebug.Location = New-Object System.Drawing.Point(300, 90)
$chkDebug.AutoSize = $true
$form.Controls.Add($chkDebug)

$txtRasSuffix = New-Object System.Windows.Forms.TextBox
$txtRasSuffix.Location = New-Object System.Drawing.Point(195,158)
$txtRasSuffix.Width = 290
$form.Controls.Add($txtRasSuffix)

# Button Run
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Выполнить"
$btnRun.Location = New-Object System.Drawing.Point(195,200)
$btnRun.Width = 150
$form.Controls.Add($btnRun)

# Logging (readonly) textbox
$tbLog = New-Object System.Windows.Forms.TextBox
$tbLog.Location = New-Object System.Drawing.Point(10,240)
$tbLog.Size = New-Object System.Drawing.Size(515,290)
$tbLog.Multiline = $true
$tbLog.ReadOnly = $true
$tbLog.ScrollBars = "Vertical"
$tbLog.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($tbLog)

# FolderBrowserDialog handlers
$browser = New-Object System.Windows.Forms.FolderBrowserDialog
$btnBrowsePlatform.Add_Click({ if($browser.ShowDialog() -eq "OK"){ $txtPlatform.Text = $browser.SelectedPath } })
$btnBrowseBaseFolder.Add_Click({ if($browser.ShowDialog() -eq "OK"){ $txtBaseFolder.Text = $browser.SelectedPath } })

# Main action on button Run
$btnRun.Add_Click({
    $PlatformPath = $txtPlatform.Text.Trim()
    $BaseFolder = $txtBaseFolder.Text.Trim()
    $StartPort = $numPort.Value.ToString()
    $RagentSuffix = $txtRagentSuffix.Text.Trim()
    $RasSuffix = $txtRasSuffix.Text.Trim()

    $ErrorActionPreference = "Stop"

    $tbLog.Clear()
    Add-Log "Путь к платформе: $PlatformPath"
    Add-Log "Верхнеуровневая папка для srvinfo: $BaseFolder"
    Add-Log "Стартовый порт: $StartPort"
    Add-Log "Приписка к DisplayName ragent: '$RagentSuffix'"
    Add-Log "Приписка к DisplayName ras: '$RasSuffix'"

    # Form srvinfo folder path automatically
    $ServiceInfo = Join-Path -Path $BaseFolder -ChildPath ("srvinfo$($StartPort)41")

    # Create srvinfo folder if not exists
    if (-not (Test-Path -Path $ServiceInfo)) {
        Add-Log "Папка '$ServiceInfo' не найдена. Создаю..."
        try {
            New-Item -ItemType Directory -Path $ServiceInfo -Force | Out-Null
            Add-Log "Папка '$ServiceInfo' успешно создана."
        } catch {
            Add-Log "ОШИБКА: Не удалось создать папку '$ServiceInfo' - $_"
            return
        }
    } else {
        Add-Log "Папка '$ServiceInfo' найдена."
    }

    # Проверки платформы
    if (-not (Test-Path "$PlatformPath\bin\ragent.exe")) {
        Add-Log "ОШИБКА: Не найден ragent.exe в папке платформы!"
        return
    }
    if (-not (Test-Path "$PlatformPath\bin\ras.exe")) {
        Add-Log "ОШИБКА: Не найден ras.exe в папке платформы!"
        return
    }

    $ServiceName = "1C:Enterprise 8.3 Server Agent (x86-64) $($StartPort)41"
    $Path = """$PlatformPath\bin\ragent.exe"""
    $ServiceInfoQuoted = """$ServiceInfo"""

    # Добавляем приписку к DisplayName ragent
    $ServiceDisplayName = "Агент сервера 1С:Предприятия 8.3 (x86-64) $($StartPort)41"
    if ($RagentSuffix -ne "") {
        $ServiceDisplayName += " $RagentSuffix"
    }

    if ($chkDebug.Checked) {
        $BinaryPath = "$Path -srvc -agent -regport $($StartPort)41 -port $($StartPort)40 -range $($StartPort)60:$($StartPort)91 -d $ServiceInfoQuoted -debug"
    } else {
        $BinaryPath = "$Path -srvc -agent -regport $($StartPort)41 -port $($StartPort)40 -range $($StartPort)60:$($StartPort)91 -d $ServiceInfoQuoted"
    }
    $CtrlPort = "$($StartPort)40"
    $AgentName = [System.Net.Dns]::GetHostEntry($env:computerName).HostName
    $RASPort = "$($StartPort)45"
    $RASPath = """$PlatformPath\bin\ras.exe"""

    $SrvcName = "1C:Enterprise 8.3 Remote Server $($StartPort)45"
    $Description = "1C:Enterprise 8.3 Remote Server $($StartPort)45"
    if ($RasSuffix -ne "") {
        $Description += " $RasSuffix"
    }

    $BinPath = "$RASPath cluster --service --port=$RASPort $AgentName`:$CtrlPort"

    try {
        $Creds = Get-Credential -Message "Учётные данные для сервиса $ServiceName"
    } catch {
        Add-Log "ОШИБКА: Не удалось получить учетные данные!"
        return
    }

    # Удаляем службы если есть
    foreach ($name in @($ServiceName, $SrvcName)) {
        if (Get-Service -Name $name -ErrorAction SilentlyContinue) {
            try {
                Add-Log "Останавливаю службу '$name'..."
                Stop-Service -Name $name -Force -ErrorAction SilentlyContinue
                Add-Log "Удаляю службу '$name'."
                sc.exe delete $name | Out-Null
            } catch {
                Add-Log "ОШИБКА: Не удалось удалить или остановить службу '$name' : $_"
            }
        }
    }

    try {
        Add-Log "Создаю службу '$SrvcName' ..."
        New-Service -Name $SrvcName -BinaryPathName $BinPath -DisplayName $Description -StartupType Automatic
        Add-Log "Запускаю службу '$SrvcName' ..."
        Start-Service -Name $SrvcName

        Add-Log "Создаю службу '$ServiceName' ..."
        New-Service -Name $ServiceName -BinaryPathName $BinaryPath -DisplayName $ServiceDisplayName -StartupType Automatic -Credential $Creds
        Add-Log "Запускаю службу '$ServiceName' ..."
        Start-Service -Name $ServiceName
    } catch {
        Add-Log "ОШИБКА: $($_.Exception.Message)"
        return
    }

    try {
        Add-Log "Останавливаю сервисы для обслуживания..."
        Stop-Service -Name $ServiceName -Force
        Stop-Service -Name $SrvcName -Force
    } catch { Add-Log "Некоторые сервисы уже остановлены."}

    Add-Log "Завершаю старые процессы rphost/rmngr..."
    foreach($process in (Get-WmiObject win32_process | Where-Object {($_.Name -eq 'rphost.exe' -or $_.Name -eq 'rmngr.exe') -and $_.CommandLine -like "$($StartPort)91*"})) {
        try {
            Stop-Process $process.ProcessId -Force
            Add-Log "Остановлен процесс ID=$($process.ProcessId) [$($process.Name)]."
        } catch {
            Add-Log "Ошибка при остановке процесса ID=$($process.ProcessId): $_"
        }
    }

    Start-Sleep -Seconds 5
    Add-Log "Копирую .lst файлы..."

    try {
        $reg_info = Get-ChildItem -Path $ServiceInfo -Force | Select-Object -ExpandProperty Fullname
        $lst_files = Get-ChildItem -Path $reg_info -Force -Recurse -Filter "*.lst" | Select-Object -ExpandProperty Fullname
        $currentTime = Get-Date -Format "yyyyMMdd_HHmmss"
        foreach ($file in $lst_files) {
            $folderPath = [System.IO.Path]::GetDirectoryName($file)
            $fileName = [System.IO.Path]::GetFileNameWithoutExtension($file)
            $fileExtension = [System.IO.Path]::GetExtension($file)
            $newFileName = $fileName + $currentTime  + ".lst_old"
            $newFilePath = Join-Path -Path $folderPath -ChildPath $newFileName
            Copy-Item -Path $file -Destination $newFilePath
            Add-Log "Скопирован '$file' -> '$newFilePath'"
        }
    } catch {
        Add-Log "ОШИБКА при копировании .lst: $_"
    }

    # Замена имени ПК на FQDN в .lst, если сервис остановлен
    $ComputerName = $env:computername
    $Domain = (Get-WmiObject Win32_ComputerSystem).domain
    $FQDN = "$ComputerName.$Domain"
    try {
        $state = (Get-WmiObject win32_service | Where-Object {$_.Name -eq $ServiceName} | Select-Object -ExpandProperty State)
        if ($state -eq "Stopped")
        {
            Add-Log "Обновляю имена в .lst файлах ($ComputerName -> $FQDN)..."
            foreach ($lst_file in $lst_files) {
                if (Test-Path $lst_file) {
                    $content = Get-Content -Path $lst_file -Raw
                    if (-not (Select-String -InputObject $content -Pattern $FQDN)) {
                        $newContent = $content -replace [regex]::Escape($ComputerName), $FQDN
                        Set-Content -Path $lst_file -Encoding UTF8 -Value $newContent
                        Add-Log "Заменено имя ПК на FQDN в '$lst_file'"
                    }
                }
            }
        }
    } catch {
        Add-Log "ОШИБКА при обновлении FQDN: $_"
    }

    # Запуск служб обратно
    try {
        Add-Log "Запускаю сервисы назад..."
        Start-Service -Name $ServiceName
        Start-Service -Name $SrvcName
        Add-Log "Готово! Все операции завершены."
    } catch {
        Add-Log "ОШИБКА при запуске служб: $_"
    }
})

$form.Topmost = $true
[void]$form.ShowDialog()
