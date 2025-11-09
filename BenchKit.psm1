function Stop-App {
    param(
        [string]$AppName
    )
    $proc = Get-Process -Name $AppName -ErrorAction SilentlyContinue
    if (-not $proc) { return }
    Write-Host "Closing application: $AppName"
    switch ($AppName) {
        "OneDrive" {
            & $proc[0].MainModule.Filename /shutdown
        }
        "Steam" {
            & $proc[0].MainModule.Filename -shutdown
        }
        "Discord" {
            taskkill.exe /im "$AppName.exe" /f
        }
        "Brave" {
            taskkill.exe /im "$AppName.exe" /f
        }
        default {
            taskkill.exe /im "$AppName.exe"
            try {
                $proc | Wait-Process -ErrorAction Stop -Timeout 3
            } catch {
                Write-Warning "Forcefully terminating remaining $AppName processes..."
                taskkill.exe /im "$AppName.exe" /f
            }
        }
    }
}

function Stop-BackgroundApps {
    param(
        [String[]]$AppsToClose = @("OneDrive", "Steam", "Discord", "Brave", "GoXLR App", "GoXLRAudioCplApp", "RTSS"),
        [String[]]$ServicesToStop = @(
            'ADPSvc','ALG','AMD Crash Defender Service','AMD External Events Utility','amd3dvcacheSvc',
            'AmdAppCompatSvc','AmdPpkgSvc','AppReadiness','AppXSvc','ApxSvc','AsusUpdateCheck',
            'autotimesvc','AxInstSV','BcastDVRUserService_*','BEService','BluetoothUserService_*',
            'brave','BraveElevationService','bravem','CDPSvc','CDPUserSvc_*','CloudBackupRestoreSvc_*',
            'DiagTrack','dmwappushservice','DoSvc','edgeupdate','edgeupdatem','EntAppSvc','fhsvc',
            'FileSyncHelper','GameInputSvc','GraphicsPerfSvc','InstallService','InventorySvc',
            'logi_lamparray_service','LxpSvc','MapsBroker','MicrosoftEdgeElevationService',
            'OneDrive Updater Service','OneSyncSvc_*','PcaSvc','PhoneSvc','PimIndexMaintenanceSvc_*',
            'PushToInstall','RasAuto','RasMan','refsdedupsvc','RemoteRegistry','Rockstar Service',
            'Spooler','Steam Client Service','SysMain','TroubleshootingSvc','UnistoreSvc_*','UsoSvc',
            'VaultSvc','VSInstallerElevationService','WaaSMedicSvc','WpnService','WpnUserService_*',
            'WSearch','wuauserv','XblAuthManager','XblGameSave','XboxGipSvc','XboxNetApiSvc'
        )
    )
    foreach ($app in $AppsToClose) {
        Stop-App -AppName $app
    }
    foreach ($svc in Get-Service) {
        foreach ($pattern in $ServicesToStop) {
            if ($svc.Name -like $pattern -and $svc.Status -eq 'Running') {
                try {
                    Write-Host "Stopping service: $($svc.Name) ($($svc.DisplayName))"
                    Stop-Service -Name $svc.Name
                } catch {
                    Write-Warning "Could not stop service: $($svc.Name) ($($svc.DisplayName))"
                }
            }
        }
    }
}

function Set-ProcessPriority {
    param(
        [Parameter(Mandatory)][String]$Name,
        [Parameter(Mandatory)][String]$Priority
    )
    Write-Host "Setting priority of $Name to $Priority..."
    $process = Get-Process $Name -ErrorAction SilentlyContinue
    if ($process) {
        try {
            $process.PriorityClass = $Priority
        } catch {
            Write-Warning "Failed to change priority: $($_.Exception.Message)"
        }
    } else {
        Write-Warning "$Name process not found. Skipping priority change."
    }
}

function Set-ProcessAffinity {
    param(
        [Parameter(Mandatory)][String]$Name,
        [Parameter(Mandatory)][Int32]$Affinity
    )
    Write-Host "Setting affinity of $Name to $Affinity..."
    $process = Get-Process $Name -ErrorAction SilentlyContinue
    if ($process) {
        try {
            $process.ProcessorAffinity = $Affinity
        } catch {
            Write-Warning "Failed to change affinity: $($_.Exception.Message)"
        }
    } else {
        Write-Warning "$Name process not found. Skipping affinity change."
    }
}

function Start-Hwinfo {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$HwinfoExe,
        [Parameter(Mandatory)][String]$Priority
    )
    $csv = Join-Path $Folder "hwinfo.csv"
    if (-not (Test-Path $Folder)) {
        Write-Host "Creating folder '$Folder'..."
        New-Item -Path $Folder -ItemType "directory" -Force
    }
    if (Test-Path $csv) {
        Write-Host "Removing existing '$csv'..."
        Remove-Item $csv -Force -ErrorAction Stop
    }
    Start-Process $HwinfoExe -WorkingDirectory $Folder
    Write-Host "Please begin sensor data logging now."
    Write-Host "- Open the HWiNFO Sensors window."
    Write-Host "- Click 'Start Logging' (spreadsheet icon)."
    Write-Host "- Save the log as: $csv"
    Write-Host "Waiting for '$csv' to appear..."
    while (-not (Test-Path $csv)) {
        Start-Sleep -Milliseconds 200
    }
    Set-ProcessPriority -Name hwinfo64 -Priority $Priority
    Write-Host "HWiNFO logging started"
}

function Stop-Hwinfo {
    param()
    while (Get-Process hwinfo64 -ErrorAction SilentlyContinue) {
        taskkill /im hwinfo64.exe
        Start-Sleep -Seconds 1
    }
    Write-Host "HWiNFO logging stopped"
}

function Invoke-Wait {
    param(
        [Parameter(Mandatory)][String]$Activity,
        [Parameter(Mandatory)][Int32]$Duration,
        [Parameter(Mandatory = $false)][ScriptBlock]$Callback
    )
    [Int32]$interval = 10
    for ($remaining = $Duration; $remaining -gt 0; $remaining -= $interval) {
        Write-Progress -Activity $Activity -Status "Running" -PercentComplete (100 - (100 * $remaining / $Duration)) -SecondsRemaining $remaining
        Start-Sleep -Seconds $interval
        if ($Callback) {
            $result = & $Callback
            if ($result -eq $false) {
                Write-Progress -Activity $Activity -Status "Aborted" -PercentComplete (100 - (100 * $remaining / $Duration)) -SecondsRemaining 0 -Completed
                return
            }
        }
    }
    Write-Progress -Activity $Activity -Status "Finished" -PercentComplete 100 -SecondsRemaining 0 -Completed
}

function Find-CinebenchValues {
    param(
        [Parameter(Mandatory)][String]$LogFile,
        [Parameter(Mandatory)][String]$OutputFile
    )
    if (-not (Test-Path $LogFile)) {
        Write-Error "Log file not found: $LogFile"
        return
    }
    $content = Get-Content $LogFile
    $line = ($content | Select-String -Pattern '^Values:' | Select-Object -Last 1).Line
    if ($line -match '\{([0-9\., ]+)\}') {
        $numbers = $matches[1] -split ',' | ForEach-Object { [Double]($_.Trim()) }
        $avg = ($numbers | Measure-Object -Average).Average
        $n = $numbers.Count
        if ($n -gt 1) {
            $variance = ($numbers | ForEach-Object { [Math]::Pow($_ - $avg, 2) } | Measure-Object -Sum).Sum / ($n - 1)
            $stddev = [Math]::Sqrt($variance)
        } else {
            Write-Warning "Need at least two runs to calculate sample standard deviation."
            $stddev = 0
        }
        $outputText = @(
            "=== $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ===",
            "Scores: $($numbers -join ', ')",
            ("Average: {0:F3}" -f $avg),
            ("Sample Standard Deviation: {0:F3}" -f $stddev)
        )
        $outputText | Out-File -FilePath $OutputFile -Encoding UTF8
    } else {
        Write-Warning "No values found in log."
    }
}

function Invoke-Cinebench {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$CinebenchExe,
        [Parameter(Mandatory)][String]$Priority,
        [Parameter(Mandatory)][Int32]$Duration,
        [Parameter(Mandatory)][String]$Arg
    )
    $log = Join-Path $Folder "cinebench.log"
    if (Test-Path $log) {
        Write-Host "Removing existing '$log'..."
        Remove-Item $log -Force -ErrorAction Stop
    }
    Write-Host "Starting Cinebench..."
    & (Join-Path $PSScriptRoot "start-cmd.bat") "$log" "$CinebenchExe" "$Arg" "g_CinebenchMinimumTestDuration=$Duration"
    Write-Host "Waiting for '$log' to appear..."
    while (-not (Test-Path $log)) {
        Start-Sleep -Milliseconds 200
    }
    Write-Host "Waiting for 'CINEBENCH AUTORUN'..."
    while (-Not ((Get-Content $log -Raw -ErrorAction Stop) -Match "CINEBENCH AUTORUN")) {
        Start-Sleep -Milliseconds 200
    }
    Set-ProcessPriority -Name Cinebench -Priority $Priority
    Invoke-Wait -Activity "Cinebench" -Duration $Duration
    Write-Host "Waiting for Cinebench to complete..."
    Wait-Process -Name Cinebench
    Find-CinebenchValues -LogFile $log -OutputFile (Join-Path $Folder "score.txt")
}

function Invoke-OCCT {
    param(
        [Parameter(Mandatory)][String]$OCCTExe,
        [Parameter(Mandatory)][String]$Priority
    )
    Write-Host "Starting OCCT..."
    & $OCCTExe
    Start-Sleep -Seconds 15
    Set-ProcessPriority -Name OCCT -Priority $Priority
    Write-Host "Waiting for OCCT to complete..."
    Wait-Process -Name OCCT
}

function Invoke-Prime95 {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$Prime95Exe,
        [Parameter(Mandatory)][String]$Priority,
        [Parameter(Mandatory)][Int32]$Duration,
        [Parameter(Mandatory)][Int32]$Cores,
        [Parameter(Mandatory)][Int32]$Affinity
    )
    $requiredLines = @(
        'CpuSupportsAVX=0',
        'CpuSupportsFMA3=0',
        'CpuSupportsFMA4=0',
        'CpuSupportsAVX2=0',
        'CpuSupportsAVX512F=0'
    )
    $primeDir = Split-Path -Parent $Prime95Exe
    $primeTxt = Join-Path $primeDir 'prime.txt'
    if (-not (Test-Path $primeTxt)) {
        Write-Error "prime.txt not found in $primeDir"
        return
    }
    $content = Get-Content $primeTxt -ErrorAction Stop
    foreach ($line in $requiredLines) {
        if (-not ($content -contains $line)) {
            $content = @($line) + $content
        }
    }
    if ($content -match '^NumCPUs=') {
        $content = $content -replace "^NumCPUs=.*", "NumCPUs=$Cores"
    } else {
        $content = @("NumCPUs=$Cores") + $content
    }
    Set-Content $primeTxt $content
    $resultsTxt = Join-Path $primeDir "results.txt"
    Remove-Item -Path $resultsTxt -ErrorAction SilentlyContinue
    Write-Host "Starting Prime95..."
    & $Prime95Exe "-t"
    Start-Sleep -Seconds 5
    Set-ProcessPriority -Name Prime95 -Priority $Priority
    Set-ProcessAffinity -Name Prime95 -Affinity $Affinity
    Invoke-Wait -Activity "Prime95" -Duration $Duration -Callback {
        if (Test-Path $resultsTxt) {
            if ((Get-Content $resultsTxt) -contains "FATAL ERROR") {
                Write-Error "Prime95 had fatal errors, system might be unstable"
                return $false
            }
        }
        return $true
    }
    taskkill.exe /im "Prime95.exe"
    Start-Sleep -Seconds 5
    if (-not (Test-Path $resultsTxt)) {
        Write-Warning "$resultsTxt not found"
    } else {
        $destPath = Join-Path $Folder "results.txt"
        Move-Item -Path $resultsTxt -Destination $destPath -Force
        Write-Host "Moved results file to $destPath"
    }
}

function Invoke-HwinfoIdle {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$HwinfoExe,
        [Parameter(Mandatory)][String]$Priority,
        [Parameter(Mandatory)][String]$Duration
    )
    Start-Hwinfo -Folder $Folder -HwinfoExe $HwinfoExe -Priority $Priority
    Invoke-Wait -Activity "Idle" -Duration $Duration
    Stop-Hwinfo
}

function Invoke-HwinfoCinebench {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$HwinfoExe,
        [Parameter(Mandatory)][String]$CinebenchExe,
        [Parameter(Mandatory)][String]$Priority,
        [Parameter(Mandatory)][String]$Duration,
        [Parameter(Mandatory)][String]$Arg
    )
    Start-Hwinfo -Folder $Folder -HwinfoExe $HwinfoExe -Priority $Priority
    Invoke-Cinebench -Folder $Folder -CinebenchExe $CinebenchExe -Priority $Priority -Duration $Duration -Arg $Arg
    Stop-Hwinfo
}

function Invoke-HwinfoOCCT {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$HwinfoExe,
        [Parameter(Mandatory)][String]$OCCTExe,
        [Parameter(Mandatory)][String]$Priority
    )
    Start-Hwinfo -Folder $Folder -HwinfoExe $HwinfoExe -Priority $Priority
    Invoke-OCCT -OCCTExe $OCCTExe -Priority $Priority
    Stop-Hwinfo
}

function Invoke-HwinfoPrime95 {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$HwinfoExe,
        [Parameter(Mandatory)][String]$Prime95Exe,
        [Parameter(Mandatory)][String]$Priority,
        [Parameter(Mandatory)][Int32]$Duration,
        [Parameter(Mandatory)][Int32]$Cores,
        [Parameter(Mandatory)][Int32]$Affinity
    )
    Start-Hwinfo -Folder $Folder -HwinfoExe $HwinfoExe -Priority $Priority
    Invoke-Prime95 -Folder $Folder -Prime95Exe $Prime95Exe -Priority $Priority -Duration $Duration -Cores $Cores -Affinity $Affinity
    Stop-Hwinfo
}

function Invoke-BenchKit {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$HwinfoExe,
        [Parameter(Mandatory)][String]$CinebenchExe,
        [Parameter(Mandatory)][String]$OCCTExe,
        [Parameter(Mandatory)][String]$Prime95Exe,
        [String]$Priority = "AboveNormal",
        [Int32]$IdleDuration = 1200,
        [Int32]$CinebenchDuration = 1200,
        [Int32]$Prime95Duration = 1200,
        [Int32]$Cores = 8
    )
    if (-not (Test-Path $Folder)) {
        Write-Host "Creating folder '$Folder'..."
        New-Item -Path $Folder -ItemType "directory" -Force
    }
    Stop-BackgroundApps
    $subfolder = Join-Path $Folder "idle"
    if (-not (Test-Path $subfolder)) {
        Invoke-HwinfoIdle -Folder $subfolder -HwinfoExe $HwinfoExe -Priority $Priority -Duration $IdleDuration
    }
    $subfolder = Join-Path $Folder "bench_cine_cpu1"
    if (-not (Test-Path $subfolder)) {
        Invoke-HwinfoCinebench -Folder $subfolder -HwinfoExe $HwinfoExe -CinebenchExe $CinebenchExe -Priority $Priority -Duration $CinebenchDuration -Arg "g_CinebenchCpu1Test=true"
    }
    $subfolder = Join-Path $Folder "bench_cine_cpux"
    if (-not (Test-Path $subfolder)) {
        Invoke-HwinfoCinebench -Folder $subfolder -HwinfoExe $HwinfoExe -CinebenchExe $CinebenchExe -Priority $Priority -Duration $CinebenchDuration -Arg "g_CinebenchCpuXTest=true"
    }
    $subfolder = Join-Path $Folder "bench_occt_cpu"
    if (-not (Test-Path $subfolder)) {
        Invoke-HwinfoOCCT -Folder $subfolder -HwinfoExe $HwinfoExe -OCCTExe $OCCTExe -Priority $Priority
    }
    $subfolder = Join-Path $Folder "bench_occt_ram"
    if (-not (Test-Path $subfolder)) {
        Invoke-HwinfoOCCT -Folder $subfolder -HwinfoExe $HwinfoExe -OCCTExe $OCCTExe -Priority $Priority
    }
    $subfolder = Join-Path $Folder "stab_occt_cpuram"
    if (-not (Test-Path $subfolder)) {
        Invoke-HwinfoOCCT -Folder $subfolder -HwinfoExe $HwinfoExe -OCCTExe $OCCTExe -Priority $Priority
    }
    $subfolder = Join-Path $Folder "stab_prime95"
    if (-not (Test-Path $subfolder)) {
        Invoke-HwinfoPrime95 -Folder $subfolder -HwinfoExe $HwinfoExe -Prime95Exe $Prime95Exe -Priority $Priority -Duration $Prime95Duration -Cores $Cores -Affinity ((1 -shl ($Cores * 2)) - 1)
    }
    for ($core = 0; $i -lt $Cores; $core++) {
        $affinity = 0x3 -shl (2 * $core)
        $subfolder = Join-Path $Folder ("stab_prime95_core{0}" -f $core)
        if (-not (Test-Path $subfolder)) {
            Invoke-HwinfoPrime95 `
                -Folder $subfolder `
                -HwinfoExe $HwinfoExe `
                -Prime95Exe $Prime95Exe `
                -Priority $Priority `
                -Duration $Prime95Duration `
                -Cores 1 `
                -Affinity $affinity
        }
    }
    Write-Host "Please reboot to restore background services"
}
