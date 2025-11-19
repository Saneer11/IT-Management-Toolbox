# ================================
# IT MANAGEMENT TOOLBOX
# ================================
Clear-Host
Write-Host "======================================="
Write-Host "          IT MANAGEMENT TOOLBOX         "
Write-Host "=======================================" -ForegroundColor Cyan

# --------------------------------
# Function 1: Cisco AnyConnect Profile Check
# --------------------------------
function CiscoProfileCheck {
    Write-Host "This program will verify if Cisco AnyConnect has the .xml file in place" -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    $Sys = Read-Host "Enter Computer Name you wish to check"

    $SourcePath = "\\shanas01-fs1\manualinstalls\Tech Scripts\Deployment\XMLFiles\alwayson-IT2.xml"
    $DestPath = "\\$Sys\C$\ProgramData\Cisco\Cisco AnyConnect Secure Mobility Client\Profile\alwayson-IT2.xml"

    function Exists {
        if (Test-Path -Path $DestPath -PathType Leaf) {
            Write-Host "$Sys has the XML file in place." -ForegroundColor Green
        }
        else {
            Write-Host "$Sys does NOT have the XML file in place." -ForegroundColor Red
            CopyXML
        }
    }

    function CopyXML {
        Write-Host "Copying the XML file to $Sys..." -ForegroundColor Yellow
        try {
            Copy-Item -Path $SourcePath -Destination ("\\$Sys\C$\ProgramData\Cisco\Cisco AnyConnect Secure Mobility Client\Profile") -Force
            Write-Host "Copy complete." -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to copy XML file. Error: $($_.Exception.Message)" -ForegroundColor Red
        }
        Exists
    }

    Exists
}

# --------------------------------
# Function 2: Trigger SCCM Client Actions
# --------------------------------
function SCCMPolicyRun {
    $PC = Read-Host "Enter the remote PC name"
    $actions = @{
        "Machine Policy Retrieval & Evaluation Cycle" = "{00000000-0000-0000-0000-000000000021}"
        "Application Deployment Evaluation Cycle"     = "{00000000-0000-0000-0000-000000000121}"
        "Software Updates Deployment Evaluation Cycle"= "{00000000-0000-0000-0000-000000000108}"
        "Discovery Data Collection Cycle"             = "{00000000-0000-0000-0000-000000000003}"
        "Hardware Inventory Cycle"                    = "{00000000-0000-0000-0000-000000000001}"
        "Software Inventory Cycle"                    = "{00000000-0000-0000-0000-000000000002}"
        "Software Metering Usage Report Cycle"        = "{00000000-0000-0000-0000-000000000031}"
    }

    foreach ($a in $actions.GetEnumerator()) {
        Write-Host "Triggering: $($a.Key) on $PC..." -ForegroundColor Cyan
        try {
            Invoke-WmiMethod -ComputerName $PC -Namespace "root\ccm" -Class "SMS_Client" -Name "TriggerSchedule" -ArgumentList $a.Value | Out-Null
            Write-Host " $($a.Key) triggered successfully." -ForegroundColor Green
        }
        catch {
            Write-Host " Failed to trigger $($a.Key): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# --------------------------------
# Function 3: Restart Remote PC
# --------------------------------
function RestartRemotePC {
    $PC = Read-Host "Enter the remote PC name or IP"
    Write-Host "Attempting to restart $PC ..." -ForegroundColor Yellow
    try {
        shutdown /r /m "\\$PC" /t 0 /f
        Write-Host "Restart command sent successfully to $PC." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to restart $PC. Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# --------------------------------
# Function 4: Compare Installed Software Between Two PCs
# --------------------------------
function CompareInstalledSoftware {
    # Ask for computer names
    $OldPC  = Read-Host "Enter the OLD PC name or IP address"
    $NewPC  = Read-Host "Enter the NEW PC name or IP address"

    Write-Host "`nCollecting installed software from both systems..." -ForegroundColor Cyan

    # Function to check WinRM connectivity
    function Test-Remote {
        param($Computer)
        try {
            Test-WsMan -ComputerName $Computer -ErrorAction Stop | Out-Null
            return $true
        } catch {
            Write-Host " Cannot connect to ${Computer} via WinRM. Please enable PS Remoting on that system." -ForegroundColor Red
            return $false
        }
    }

    # Function to collect installed software
    function Get-InstalledSoftware {
        param($ComputerName)

        if (-not (Test-Remote -Computer $ComputerName)) {
            return @()
        }

        try {
            Write-Host "→ Querying ${ComputerName}..." -ForegroundColor Yellow
            $result = Invoke-Command -ComputerName $ComputerName -ErrorAction Stop -ScriptBlock {
                $paths = @(
                    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                    'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
                )
                $software = foreach ($path in $paths) {
                    Get-ChildItem -Path $path -ErrorAction SilentlyContinue | ForEach-Object {
                        $p = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
                        if ($p.DisplayName) {
                            [PSCustomObject]@{
                                Name    = $p.DisplayName
                                Version = $p.DisplayVersion
                            }
                        }
                    }
                }
                $software | Sort-Object Name -Unique
            }
            return $result
        } catch {
            Write-Host " Failed to query ${ComputerName}: $($_.Exception.Message)" -ForegroundColor Red
            return @()
        }
    }

    # Get software lists
    $OldList = Get-InstalledSoftware -ComputerName $OldPC
    $NewList = Get-InstalledSoftware -ComputerName $NewPC

    # Validate data
    if (-not $OldList -or $OldList.Count -eq 0) {
        Write-Host "`nNo software retrieved from ${OldPC}. Please verify WinRM access and permissions." -ForegroundColor Red
        return
    }
    if (-not $NewList -or $NewList.Count -eq 0) {
        Write-Host "`nNo software retrieved from ${NewPC}. Please verify WinRM access and permissions." -ForegroundColor Red
        return
    }

    # Compare by software name
    $OldNames = $OldList.Name | Sort-Object -Unique
    $NewNames = $NewList.Name | Sort-Object -Unique

    $Missing = $OldNames | Where-Object { $_ -notin $NewNames }

    Write-Host "`nSoftware installed on ${OldPC} but missing on ${NewPC}:`n" -ForegroundColor Green

    if ($Missing -and $Missing.Count -gt 0) {
        $Missing | Sort-Object | Format-Table -AutoSize

        # Optional export
        $save = Read-Host "`nDo you want to export the missing list to CSV? (Y/N)"
        if ($save -match '^[Yy]$') {
            $path = Read-Host "Enter full path to save CSV file (e.g. C:\Temp\MissingSoftware.csv)"
            $Missing | ForEach-Object { [PSCustomObject]@{ Software = $_ } } |
                Export-Csv -Path $path -NoTypeInformation -Force
            Write-Host " Export completed: $path" -ForegroundColor Yellow
        }
    } else {
        Write-Host " All software from ${OldPC} already exists on ${NewPC}." -ForegroundColor Green
    }
}


# --------------------------------
# Main Menu
# --------------------------------
do {
    Write-Host ""
    Write-Host "Select an option:"
    Write-Host "1. Cisco AnyConnect Profile Check"
    Write-Host "2. Trigger SCCM Client Policy Actions"
    Write-Host "3. Restart Remote PC"
    Write-Host "4. Compare Installed Software Between Two PCs"
    Write-Host "0. Exit"
    $choice = Read-Host "Enter your choice"

    switch ($choice) {
        "1" { CiscoProfileCheck }
        "2" { SCCMPolicyRun }
        "3" { RestartRemotePC }
        "4" { CompareInstalledSoftware }
        "0" { Write-Host "Exiting Toolbox..." -ForegroundColor Yellow }
        default { Write-Host "Invalid choice. Please try again." -ForegroundColor Red }
    }

} while ($choice -ne "0")
