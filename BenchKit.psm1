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
            Write-Host "Killing $AppName..."
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

function Start-Hwinfo {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$HwinfoExe
    )
    $csv = "$Folder\hwinfo.csv"
    if (-not (Test-Path $Folder)) {
        Write-Host "Creating folder '$Folder'..."
        New-Item -Path "." -Name $Folder -ItemType "directory"
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
    taskkill /im hwinfo64.exe
    Get-Process hwinfo64 | Wait-Process -ErrorAction Stop -Timeout 3
    Write-Host "HWiNFO logging stopped"
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
    Write-Host "Setting priority to $Priority..."
    $process = Get-Process Cinebench -ErrorAction SilentlyContinue
    if ($process) {
        try {
            $process.PriorityClass = $Priority
        } catch {
            Write-Host "Failed to change priority: $($_.Exception.Message)"
        }
    } else {
        Write-Warning "Cinebench process not found. Skipping priority change."
    }
    Write-Host "Waiting for Cinebench to complete..."
    Wait-Process -Name Cinebench
}

function Invoke-OCCT {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$OCCTExe,
        [Parameter(Mandatory)][String]$Priority
    )
    Write-Host "Starting OCCT..."
    & $OCCTExe
    Write-Host "Setting priority to $Priority..."
    $process = Get-Process OCCT -ErrorAction SilentlyContinue
    if ($process) {
        try {
            $process.PriorityClass = $Priority
        } catch {
            Write-Host "Failed to change priority: $($_.Exception.Message)"
        }
    } else {
        Write-Warning "OCCT process not found. Skipping priority change."
    }
    Write-Host "Waiting for OCCT to complete..."
    Wait-Process -Name OCCT
}

function Invoke-HwinfoIdle {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$HwinfoExe,
        [Parameter(Mandatory)][String]$Duration
    )
    Start-Hwinfo -Folder $Folder -hwinfoExe $HwinfoExe
    $interval = 10 # update every 10 seconds
    for ($remaining = $Duration; $remaining -ge 0; $remaining -= $interval) {
        Write-Host "Time remaining: $remaining"
        Start-Sleep -Seconds $interval
    }
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
    Start-Hwinfo -Folder $Folder -hwinfoExe $HwinfoExe
    Invoke-Cinebench -Folder $Folder -CinebenchExe $CinebenchExe -Priority $Priority -Duration $Duration
    Stop-Hwinfo
}

function Invoke-HwinfoOCCT {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$HwinfoExe,
        [Parameter(Mandatory)][String]$OCCTExe,
        [Parameter(Mandatory)][String]$Priority,
        [Parameter(Mandatory)][String]$Duration
    )
    Start-Hwinfo -Folder $Folder -hwinfoExe $HwinfoExe
    Invoke-OCCT -Folder $Folder -OCCTExe $OCCTExe -Priority $Priority
    Stop-Hwinfo
}

function Invoke-BenchKit {
    param(
        [Parameter(Mandatory)][String]$Folder,
        [Parameter(Mandatory)][String]$HwinfoExe,
        [Parameter(Mandatory)][String]$CinebenchExe,
        [Parameter(Mandatory)][String]$OCCTExe,
        [String]$Priority = "High",
        [Int32]$IdleDuration = 600,
        [Int32]$CinebenchDuration = 900
    )
    Stop-BackgroundApps
    if (-not (Test-Path $Folder)) {
        Write-Host "Creating folder '$Folder'..."
        New-Item -Path "." -Name $Folder -ItemType "directory"
    }
    Invoke-HwinfoIdle -Folder "$Folder\idle" -HwinfoExe $HwinfoExe -Duration $IdleDuration
    Invoke-HwinfoCinebench -Folder "$Folder\bench_cine_cpu" -HwinfoExe $HwinfoExe -CinebenchExe $CinebenchExe -Priority $Priority -Duration $CinebenchDuration
    Invoke-HwinfoOCCT -Folder "$Folder\stab_occt_cpu" -HwinfoExe $HwinfoExe -OCCTExe $OCCTExe -Priority $Priority
    Invoke-HwinfoOCCT -Folder "$Folder\bench_occt_cpu" -HwinfoExe $HwinfoExe -OCCTExe $OCCTExe -Priority $Priority
    Invoke-HwinfoOCCT -Folder "$Folder\bench_occt_ram" -HwinfoExe $HwinfoExe -OCCTExe $OCCTExe -Priority $Priority
}
