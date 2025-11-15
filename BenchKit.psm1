function Get-AverageStderr {
    param(
        [Double[]]$Numbers = @()
    )
    $n = $Numbers.Count
    $mean =
        if ($n -Eq 0) {
            [Double]::NaN
        } else {
            [Double](($Numbers | Measure-Object -Average).Average)
        }
    $stderr =
        if ($n -Lt 2) {
            [Double]::PositiveInfinity
        } else {
            $sumSqDiff = ($Numbers | ForEach-Object { ($_ - $mean) * ($_ - $mean) } | Measure-Object -Sum).Sum
            $variance = $sumSqDiff / ($n - 1)
            [Math]::Sqrt($variance / $n)
        }
    [PSCustomObject]@{
        Average = $mean
        StdErr  = $stderr
    }
}

function Get-ConfidenceInterval {
    param(
        [Double[]]$Numbers = @()
    )
    $avgstderr = Get-AverageStderr -Numbers $Numbers
    if ([Double]::IsPositiveInfinity($avgstderr.StdErr)) {
        @(-[Double]::PositiveInfinity, [Double]::PositiveInfinity)
    } else {
        @(
            [Double]($avgstderr.Average - 1.96 * $avgstderr.StdErr),
            [Double]($avgstderr.Average + 1.96 * $avgstderr.StdErr)
        )
    }
}

function Get-FormattedNumbers {
    param(
        [Parameter(Mandatory)][String]$Name,
        [Double[]]$Numbers = @(),
        [String]$Format = "F3"
    )
    $confint = Get-ConfidenceInterval -Numbers $Numbers
    Write-Output "=== $Name ==="
    Write-Output (($Numbers | ForEach-Object { "{0:$Format}" -f $_ }) -join ",")
    Write-Output ("95% CI for sample mean: [{0:$Format}, {1:$Format}]" -f $confint)
}

function Get-FormattedDate {
    Write-Output "=== Date ==="
    Write-Output (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
}

function Stop-App {
    param(
        [String]$AppName
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
        [String[]]$AppsToClose = @("OneDrive", "Discord", "Brave", "GoXLR App", "GoXLRAudioCplApp", "RTSS"),
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
    $processes = Get-Process $Name -ErrorAction SilentlyContinue
    foreach ($process in $processes) {
        try {
            $process.PriorityClass = $Priority
            Write-Host "Priority of $process set to $Priority"
        } catch {
            Write-Warning "Failed to change priority on $($process): $($_.Exception.Message)"
        }
    }
}

function Set-ProcessAffinity {
    param(
        [Parameter(Mandatory)][String]$Name,
        [Parameter(Mandatory)][Int32]$Affinity
    )
    Write-Host "Setting affinity of $Name to $Affinity..."
    $processes = Get-Process $Name -ErrorAction SilentlyContinue
    foreach ($process in $processes) {
        try {
            $process.ProcessorAffinity = $Affinity
            Write-Host "Affinity of $process set to $Affinity"
        } catch {
            Write-Warning "Failed to change affinity on $($process): $($_.Exception.Message)"
        }
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
        $confint = Get-ConfidenceInterval -Numbers $numbers
        $outputText = @(
            (Get-FormattedDate),
            (Get-FormattedNumbers -Name "Score" -Numbers $numbers)
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

function Invoke-Cyberpunk {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$CyberpunkExe,
        [Parameter(Mandatory)][String]$Priority,
        [Int32]$Runs = 5
    )
    $resultsrootfolder = Join-Path $([Environment]::GetFolderPath('MyDocuments')) "CD Projekt Red\Cyberpunk 2077\benchmarkResults"
    if (Test-Path $resultsrootfolder) {
        Write-Host "Removing old benchmark results folder: $resultsrootfolder"
        Remove-Item $resultsrootfolder -Recurse -Force
    }
    ForEach ($run in @(1..$Runs)) {
        Write-Host "Starting Cyberpunk (run $run/$Runs)..."
        & "$CyberpunkExe" "-benchmark"
        Start-Sleep -Seconds 10
        Set-ProcessPriority -Name Cyberpunk2077 -Priority $Priority
        Write-Host "Waiting for Cyberpunk to complete..."
        Wait-Process -Name Cyberpunk2077
    }
    $i=0; Get-ChildItem -Path $resultsrootfolder "summary.json" -Recurse | ForEach-Object { $i++; Copy-Item $_.FullName (Join-Path $Folder "summary$i.json") }
    $allFps =  Get-ChildItem -File -Path $resultsrootfolder -Filter "summary.json" -Recurse |
    ForEach-Object { 
        try { 
            ConvertFrom-Json (Get-Content $_.FullName -Raw) 
        } catch {
            Write-Warning "Skipping $_"
        }
    } |
    ForEach-Object { $_.Data } |
    Where-Object { $_ } |
    ForEach-Object {
        [PSCustomObject]@{
            Avg = $_.averageFps
            Min = $_.minFps
            Max = $_.maxFps
        }
    }
    $outputText = @(
        (Get-FormattedDate),
        (Get-FormattedNumbers -Name "AvgFps" -Numbers ($allFps | Where-Object Avg | ForEach-Object Avg)),
        (Get-FormattedNumbers -Name "MinFps" -Numbers ($allFps | Where-Object Min | ForEach-Object Min)),
        (Get-FormattedNumbers -Name "MaxFps" -Numbers ($allFps | Where-Object Max | ForEach-Object Max))
    )
    $outputText | Out-File -FilePath (Join-Path $Folder "fps.txt") -Encoding UTF8
}

function Invoke-3DMark {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$Priority
    )
    $resultsrootfolder = Join-Path $([Environment]::GetFolderPath('MyDocuments')) "3DMark"
    if (Test-Path $resultsrootfolder) {
        Write-Host "Removing old benchmark results folder: $resultsrootfolder"
        Remove-Item $resultsrootfolder -Recurse -Force
    }
    Write-Host "Starting 3DMark..."
    # appid for 3dmark demo is 231350 (does not run without steam)
    Start-Process "steam://rungameid/231350"
    # application takes a while to initialize
    Start-Sleep -Seconds 45
    Set-ProcessPriority -Name 3DMark -Priority $Priority
    Write-Host "Please run the 3DMark Steel Nomad test five times, then exit the application..."
    Wait-Process -Name 3DMark
    if (Test-Path $resultsrootfolder) {
        $scores = Get-ChildItem -File -Path $resultsrootfolder -Filter "*.3dmark-result" -Recurse |
            ForEach-Object {
                if ($_.BaseName -match '^3DMark-SteelNomad-([0-9]+)-') {
                    $matches[1]
                } else {
                    write-Error "Result file does not match expected pattern: $($_.BaseName)"
                }
            }
        $outputText = @(
            (Get-FormattedDate),
            (Get-FormattedNumbers -Name "Score" -Numbers $scores)
        )
        $outputText | Out-File -FilePath (Join-Path $Folder "score.txt") -Encoding UTF8
    } else {
        Write-Error "Benchmark results folder not found: $resultsrootfolder"
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

function Invoke-HwinfoCyberpunk {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$HwinfoExe,
        [Parameter(Mandatory)][String]$CyberpunkExe,
        [Parameter(Mandatory)][String]$Priority
    )
    Start-Hwinfo -Folder $Folder -HwinfoExe $HwinfoExe -Priority $Priority
    Invoke-Cyberpunk -Folder $Folder -CyberpunkExe $CyberpunkExe -Priority $Priority
    Stop-Hwinfo
}

function Invoke-Hwinfo3DMark {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$HwinfoExe,
        [Parameter(Mandatory)][String]$Priority
    )
    Start-Hwinfo -Folder $Folder -HwinfoExe $HwinfoExe -Priority $Priority
    Invoke-3DMark -Folder $Folder -Priority $Priority
    Stop-Hwinfo
}

function Invoke-BenchKit {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$HwinfoExe,
        [Parameter(Mandatory)][String]$CinebenchExe,
        [Parameter(Mandatory)][String]$OCCTExe,
        [Parameter(Mandatory)][String]$Prime95Exe,
        [Parameter(Mandatory)][String]$CyberpunkExe,
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
    $jobs = @(
        @{
            Name = "idle"
            Script = {
                param($path)
                Invoke-HwinfoIdle -Folder $path -HwinfoExe $HwinfoExe -Priority $Priority -Duration $IdleDuration
            }
        }

        @{
            Name = "bench_cyberpunk"
            Script = {
                param($path)
                Invoke-HwinfoCyberpunk -Folder $path -HwinfoExe $HwinfoExe -CyberpunkExe $CyberpunkExe -Priority $Priority
            }
        }
        @{
            Name = "bench_cine_cpu1"
            Script = {
                param($path)
                Invoke-HwinfoCinebench -Folder $path -HwinfoExe $HwinfoExe -CinebenchExe $CinebenchExe -Priority $Priority -Duration $CinebenchDuration -Arg "g_CinebenchCpu1Test=true"
            }
        }
        @{
            Name = "bench_cine_cpux"
            Script = {
                param($path)
                Invoke-HwinfoCinebench -Folder $path -HwinfoExe $HwinfoExe -CinebenchExe $CinebenchExe -Priority $Priority -Duration $CinebenchDuration -Arg "g_CinebenchCpuXTest=true"
            }
        }
        @{
            Name = "bench_3dmark_steelnomad"
            Script = {
                param($path)
                Invoke-Hwinfo3DMark -Folder $path -HwinfoExe $HwinfoExe -Priority $Priority
            }
        }
        "bench_occt_cpu","bench_occt_ram","stab_occt_cpuram","stab_occt_3dvar","stab_occt_3dswitch","stab_occt_vram" | ForEach-Object {
            @{
                Name = $_
                Script = {
                    param($path)
                    Invoke-HwinfoOCCT -Folder $path -HwinfoExe $HwinfoExe -OCCTExe $OCCTExe -Priority $Priority
                }
            }
        }
        @{
            Name = "stab_prime95"
            Script = {
                param($path)
                $affinity = ((1 -shl ($Cores * 2)) - 1)
                Invoke-HwinfoPrime95 -Folder $path -HwinfoExe $HwinfoExe -Prime95Exe $Prime95Exe -Priority $Priority -Duration $Prime95Duration -Cores $Cores -Affinity $affinity
            }
        }
        0..($Cores - 1) | ForEach-Object {
            $core = $_
            @{
                Name = "stab_prime95_core$core"
                Script = {
                    param($path)
                    $affinity = (0x3 -shl (2 * $core))
                    Invoke-HwinfoPrime95 -Folder $path -HwinfoExe $HwinfoExe -Prime95Exe $Prime95Exe -Priority $Priority -Duration $Prime95Duration -Cores 1 -Affinity $affinity
                }
            }
        }
    )
    foreach ($job in $jobs) {
        $path = Join-Path $Folder $job.Name
        if (-not (Test-Path $path)) {
            Write-Host "Running job '$($job.Name)'..."
            & $job.Script $path
        }
    }
    Write-Host "Please reboot to restore background services"
}
