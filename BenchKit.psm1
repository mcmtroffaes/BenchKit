function Stop-App {
    param(
        [string]$AppName
    )
    $proc = Get-Process -Name $AppName -ErrorAction SilentlyContinue
    if (-not $proc) { return }
    Write-Host "Closing $AppName..."
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
        [String[]]$AppsToClose = @("OneDrive", "Steam", "Discord", "Brave", "GoXLR App", "GoXLRAudioCplApp", "RTSS")
    )

    Write-Host "Checking for running background apps..."
    foreach ($app in $AppsToClose) {
        Stop-App -AppName $app
    }
    Write-Host "Background apps closed."
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

function Start-Hwinfo {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$HwinfoExe
    )
    $csv = "$Folder\hwinfo.csv"
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
        [Parameter(Mandatory)][Int32]$Duration
    )
    [Int32]$interval = 10
    for ($remaining = $Duration; $remaining -gt 0; $remaining -= $interval) {
        Write-Progress -Activity "Idle" -Status "Waiting" -PercentComplete (100 - (100 * $remaining / $Duration)) -SecondsRemaining $remaining
        Start-Sleep -Seconds $interval
    }
}

function Invoke-Cinebench {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$CinebenchExe,
        [Parameter(Mandatory)][String]$Priority,
        [Parameter(Mandatory)][Int32]$Duration
    )
    $log = "$Folder\cinebench.log"
    if (Test-Path $log) {
        Write-Host "Removing existing '$log'..."
        Remove-Item $log -Force -ErrorAction Stop
    }
    Write-Host "Starting Cinebench..."
    & $PSScriptRoot\start-cmd.bat "$log" "$CinebenchExe" "g_CinebenchCpuXTest=true" "g_CinebenchMinimumTestDuration=$Duration"
    Write-Host "Waiting for '$log' to appear..."
    while (-not (Test-Path $log)) {
        Start-Sleep -Milliseconds 200
    }
    Write-Host "Waiting for 'CINEBENCH AUTORUN'..."
    while (-Not ((Get-Content $log -Raw -ErrorAction Stop) -Match "CINEBENCH AUTORUN")) {
        Start-Sleep -Milliseconds 200
    }
    Set-ProcessPriority -Name Cinebench -Priority $Priority
    Write-Host "Waiting for Cinebench to complete..."
    Wait-Process -Name Cinebench
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
        [Parameter(Mandatory)][String]$Prime95Exe,
        [Parameter(Mandatory)][String]$Priority,
        [Parameter(Mandatory)][Int32]$Duration
    )
    Write-Host "Starting Prime95..."
    & $Prime95Exe -t8
    Start-Sleep -Seconds 5
    Set-ProcessPriority -Name Prime95 -Priority $Priority
    Invoke-Wait -Duration $Duration
    taskkill.exe /im "Prime95.exe"
}

function Invoke-HwinfoIdle {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$HwinfoExe,
        [Parameter(Mandatory)][String]$Duration
    )
    Start-Hwinfo -Folder $Folder -HwinfoExe $HwinfoExe
    Invoke-Wait -Duration $Duration
    Stop-Hwinfo
}

function Invoke-HwinfoCinebench {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$HwinfoExe,
        [Parameter(Mandatory)][String]$CinebenchExe,
        [Parameter(Mandatory)][String]$Priority,
        [Parameter(Mandatory)][String]$Duration
    )
    Start-Hwinfo -Folder $Folder -HwinfoExe $HwinfoExe
    Invoke-Cinebench -Folder $Folder -CinebenchExe $CinebenchExe -Priority $Priority -Duration $Duration
    Stop-Hwinfo
}

function Invoke-HwinfoOCCT {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$HwinfoExe,
        [Parameter(Mandatory)][String]$OCCTExe,
        [Parameter(Mandatory)][String]$Priority
    )
    Start-Hwinfo -Folder $Folder -HwinfoExe $HwinfoExe
    Invoke-OCCT -OCCTExe $OCCTExe -Priority $Priority
    Stop-Hwinfo
}

function Invoke-HwinfoPrime95 {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$HwinfoExe,
        [Parameter(Mandatory)][String]$Prime95Exe,
        [Parameter(Mandatory)][String]$Priority,
        [Parameter(Mandatory)][Int32]$Duration
    )
    Start-Hwinfo -Folder $Folder -HwinfoExe $HwinfoExe
    Invoke-Prime95 -Prime95Exe $Prime95Exe -Priority $Priority -Duration $Duration
    Stop-Hwinfo
}

function Invoke-BenchKit {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$HwinfoExe,
        [Parameter(Mandatory)][String]$CinebenchExe,
        [Parameter(Mandatory)][String]$OCCTExe,
        [Parameter(Mandatory)][String]$Prime95Exe,
        [String]$Priority = "High",
        [Int32]$IdleDuration = 600,
        [Int32]$CinebenchDuration = 900,
        [Int32]$Prime95Duration = 600
    )
    Stop-BackgroundApps
    if (-not (Test-Path $Folder)) {
        Write-Host "Creating folder '$Folder'..."
        New-Item -Path $Folder -ItemType "directory" -Force
    }
    Invoke-HwinfoIdle -Folder "$Folder\idle" -HwinfoExe $HwinfoExe -Duration $IdleDuration
    Invoke-HwinfoCinebench -Folder "$Folder\bench_cine_cpu" -HwinfoExe $HwinfoExe -CinebenchExe $CinebenchExe -Priority $Priority -Duration $CinebenchDuration
    Invoke-HwinfoOCCT -Folder "$Folder\bench_occt_cpu" -HwinfoExe $HwinfoExe -OCCTExe $OCCTExe -Priority $Priority
    Invoke-HwinfoOCCT -Folder "$Folder\bench_occt_ram" -HwinfoExe $HwinfoExe -OCCTExe $OCCTExe -Priority $Priority
    Invoke-HwinfoOCCT -Folder "$Folder\stab_occt_cpu" -HwinfoExe $HwinfoExe -OCCTExe $OCCTExe -Priority $Priority
    Invoke-HwinfoPrime95 -Folder "$Folder\stab_prime95" -HwinfoExe $HwinfoExe -Prime95Exe $Prime95Exe -Priority $Priority -Duration $Prime95Duration
}
