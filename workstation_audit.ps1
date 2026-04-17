# ============================================================
# UNIVERSAL WORKSTATION AUDIT SCRIPT
#
# Script Version : 2.0.0
# Schema Version : 2.0      (output JSON shape; bump on breaking changes)
# Build Date     : 2026-04-17
# Sections       : 49       (originals 1-33; security/diag additions 34-49)
# Compatibility  : Windows 10 21H2 -> Windows 11 25H2 (admin required)
#
# Outputs        : C:\Temp\<HOSTNAME>_workstation_audit_<DATE>.json
#                  (UTF-8, ConvertTo-Json depth 10)
#
# Each top-level JSON key is one of: _metadata, _errors, then one
# property per section in execution order. Section names are stable;
# never rename without bumping Schema Version.
# ============================================================

$ScriptVersion       = "2.0.0"
$ScriptSchemaVersion = "2.0"
$ScriptBuildDate     = "2026-04-17"

# Auto-relaunch with ExecutionPolicy Bypass if needed
if ($MyInvocation.MyCommand.Path) {
    $currentPolicy = Get-ExecutionPolicy -Scope Process
    if ($currentPolicy -eq 'Restricted' -or $currentPolicy -eq 'AllSigned') {
        Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File `"$($MyInvocation.MyCommand.Path)`"" -Wait -NoNewWindow
        exit
    }
}

# Verify running as admin
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "ERROR: This script must be run as Administrator." -ForegroundColor Red
    exit 1
}

$Date = Get-Date -Format 'yyyy-MM-dd'
$OutputDir = "C:\Temp"
$Name = $env:COMPUTERNAME
$JsonFile = "$OutputDir\${Name}_workstation_audit_$Date.json"

if (!(Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }

# Self-SHA (file-on-disk runs only; iex-streamed runs report "iex-streamed")
$ScriptSelfSHA256 = if ($MyInvocation.MyCommand.Path -and (Test-Path $MyInvocation.MyCommand.Path)) {
    try { (Get-FileHash -Path $MyInvocation.MyCommand.Path -Algorithm SHA256 -ErrorAction Stop).Hash }
    catch { "sha-computation-failed" }
} else { "iex-streamed" }

# Structured data collector
$audit = [ordered]@{
    _metadata = [ordered]@{
        ScriptVersion       = $ScriptVersion
        ScriptSchemaVersion = $ScriptSchemaVersion
        ScriptBuildDate     = $ScriptBuildDate
        ScriptSelfSHA256    = $ScriptSelfSHA256
        ScriptType          = "Workstation"
        SectionCount        = 49
        RunDate             = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        RunDateUtc          = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        RunBy               = "$env:USERDOMAIN\$env:USERNAME"
        Hostname            = $env:COMPUTERNAME
        OSCaption           = (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption
        OSBuild             = (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue).BuildNumber
        PowerShellVersion   = "$($PSVersionTable.PSVersion)"
    }
    _errors = @()
}

Write-Host "============================================="
Write-Host "  UNIVERSAL WORKSTATION AUDIT v$ScriptVersion (schema $ScriptSchemaVersion)"
Write-Host "  Build $ScriptBuildDate  -  $($audit._metadata.SectionCount) sections"
Write-Host "  $($audit._metadata.RunDate)"
Write-Host "============================================="
Write-Host ""

# =====================================================
# 1. SYSTEM INFO
# =====================================================
Write-Host "=== 1. SYSTEM INFO ===" -ForegroundColor Cyan
Write-Host "  Running: Get-CimInstance Win32_OperatingSystem, Win32_ComputerSystem"
try {
    $os = Get-CimInstance Win32_OperatingSystem
    $cs = Get-CimInstance Win32_ComputerSystem
    $uptime = (Get-Date) - $os.LastBootUpTime

    # Friendly version (e.g., 23H2)
    $displayVersion = "N/A"
    try {
        $displayVersion = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop).DisplayVersion
    } catch {}

    $sysInfo = [ordered]@{
        Hostname       = $env:COMPUTERNAME
        OS             = $os.Caption
        OSVersion      = $os.Version
        DisplayVersion = $displayVersion
        BuildNumber    = $os.BuildNumber
        Architecture   = $os.OSArchitecture
        InstallDate    = $os.InstallDate.ToString("yyyy-MM-dd")
        LastBoot       = $os.LastBootUpTime.ToString("yyyy-MM-dd HH:mm:ss")
        UptimeDays     = [math]::Round($uptime.TotalDays, 1)
        Domain         = $cs.Domain
        DomainJoined   = $cs.PartOfDomain
        CurrentUser    = $cs.UserName
        TotalRAM_GB    = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
    }

    $sysInfo.GetEnumerator() | ForEach-Object { Write-Host "  $($_.Key): $($_.Value)" }
    $audit.SystemInfo = $sysInfo
    Write-Host "  [OK] System info collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "SystemInfo"; Error = $_.Exception.Message }
}

# =====================================================
# 2. HARDWARE
# =====================================================
Write-Host ""
Write-Host "=== 2. HARDWARE ===" -ForegroundColor Cyan
Write-Host "  Running: Get-CimInstance Win32_ComputerSystem, Win32_Processor, Win32_BIOS, Win32_SystemEnclosure"
try {
    $hw = Get-CimInstance Win32_ComputerSystem
    $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
    $bios = Get-CimInstance Win32_BIOS
    $encl = Get-CimInstance Win32_SystemEnclosure | Select-Object -First 1
    $battery = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue

    $hardware = [ordered]@{
        Manufacturer = $hw.Manufacturer
        Model        = $hw.Model
        SerialNumber = if ($bios.SerialNumber) { $bios.SerialNumber } else { $encl.SerialNumber }
        CPU          = $cpu.Name
        CPUCores     = $cpu.NumberOfCores
        CPULogical   = $cpu.NumberOfLogicalProcessors
        RAM_GB       = [math]::Round($hw.TotalPhysicalMemory / 1GB, 1)
        BIOSVersion  = $bios.SMBIOSBIOSVersion
        ChassisType  = switch ($encl.ChassisTypes[0]) {
            3 {"Desktop"} 4 {"Low Profile Desktop"} 5 {"Pizza Box"} 6 {"Mini Tower"}
            7 {"Tower"} 8 {"Portable"} 9 {"Laptop"} 10 {"Notebook"} 11 {"Hand Held"}
            12 {"Docking Station"} 13 {"All in One"} 14 {"Sub Notebook"} 15 {"Space-Saving"}
            default {"Other ($($encl.ChassisTypes[0]))"}
        }
        HasBattery   = ($null -ne $battery)
    }

    $hardware.GetEnumerator() | ForEach-Object { Write-Host "  $($_.Key): $($_.Value)" }
    $audit.Hardware = $hardware
    Write-Host "  [OK] Hardware info collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "Hardware"; Error = $_.Exception.Message }
}

# =====================================================
# 3. OS ACTIVATION STATUS
# =====================================================
Write-Host ""
Write-Host "=== 3. OS ACTIVATION STATUS ===" -ForegroundColor Cyan
Write-Host "  Running: Get-CimInstance SoftwareLicensingProduct"
try {
    $lic = Get-CimInstance SoftwareLicensingProduct |
        Where-Object { $_.PartialProductKey -and $_.Name -like '*Windows*' } |
        Select-Object -First 1

    $statusText = switch ($lic.LicenseStatus) {
        0 {"Unlicensed"} 1 {"Licensed"} 2 {"OOB Grace"} 3 {"OOT Grace"}
        4 {"Non-Genuine"} 5 {"Notification"} 6 {"Extended Grace"} default {"Unknown"}
    }

    $activation = [ordered]@{
        ProductName    = $lic.Name
        LicenseStatus  = $statusText
        StatusCode     = $lic.LicenseStatus
        PartialKey     = $lic.PartialProductKey
        LicenseFamily  = $lic.LicenseFamily
    }

    $activation.GetEnumerator() | ForEach-Object { Write-Host "  $($_.Key): $($_.Value)" }

    if ($lic.LicenseStatus -ne 1) {
        Write-Host "  WARNING: This machine is NOT fully licensed!" -ForegroundColor Red
    }

    $audit.Activation = $activation
    Write-Host "  [OK] Activation status collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "Activation"; Error = $_.Exception.Message }
}

# =====================================================
# 4. STORAGE
# =====================================================
Write-Host ""
Write-Host "=== 4. STORAGE ===" -ForegroundColor Cyan
Write-Host "  Running: Get-Volume, Get-PhysicalDisk"
try {
    $volumes = Get-Volume | Where-Object { $_.DriveLetter } | Select-Object DriveLetter,
        FileSystemLabel, FileSystem,
        @{N='SizeGB';E={[math]::Round($_.Size/1GB,1)}},
        @{N='FreeGB';E={[math]::Round($_.SizeRemaining/1GB,1)}},
        @{N='UsedPct';E={if($_.Size -gt 0){[math]::Round(($_.Size - $_.SizeRemaining)/$_.Size * 100,0)}else{0}}},
        HealthStatus
    $volumes | Format-Table -AutoSize
    $audit.Volumes = @($volumes | ForEach-Object {
        [ordered]@{
            Drive = "$($_.DriveLetter):"; Label = $_.FileSystemLabel; FileSystem = $_.FileSystem
            SizeGB = $_.SizeGB; FreeGB = $_.FreeGB; UsedPct = $_.UsedPct; Health = "$($_.HealthStatus)"
        }
    })
    Write-Host "  [OK] Volumes collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "Volumes"; Error = $_.Exception.Message }
}

Write-Host "  Running: Get-PhysicalDisk"
try {
    $disks = Get-PhysicalDisk | Select-Object DeviceId, FriendlyName, MediaType,
        @{N='SizeGB';E={[math]::Round($_.Size/1GB,1)}}, HealthStatus, OperationalStatus
    $disks | Format-Table -AutoSize
    $audit.PhysicalDisks = @($disks | ForEach-Object {
        [ordered]@{
            DeviceId = "$($_.DeviceId)"; Name = $_.FriendlyName; MediaType = "$($_.MediaType)"
            SizeGB = $_.SizeGB; Health = "$($_.HealthStatus)"; Status = "$($_.OperationalStatus)"
        }
    })
    Write-Host "  [OK] Physical disks collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] PhysicalDisk: $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "PhysicalDisks"; Error = $_.Exception.Message }
}

# =====================================================
# 5. BITLOCKER STATUS
# =====================================================
Write-Host ""
Write-Host "=== 5. BITLOCKER STATUS ===" -ForegroundColor Cyan
Write-Host "  Running: Get-BitLockerVolume"
try {
    $bl = Get-BitLockerVolume -ErrorAction Stop
    $bl | ForEach-Object {
        Write-Host "  $($_.MountPoint): $($_.VolumeStatus), Protection: $($_.ProtectionStatus), Method: $($_.EncryptionMethod)"
    }
    $audit.BitLocker = @($bl | ForEach-Object {
        [ordered]@{
            MountPoint = "$($_.MountPoint)"; VolumeStatus = "$($_.VolumeStatus)"
            Protection = "$($_.ProtectionStatus)"; Method = "$($_.EncryptionMethod)"
            KeyProtectors = @($_.KeyProtector | ForEach-Object { "$($_.KeyProtectorType)" })
        }
    })
    Write-Host "  [OK] BitLocker status collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Trying manage-bde fallback..." -ForegroundColor Yellow
    try {
        $bde = manage-bde -status 2>&1
        $bde | ForEach-Object { Write-Host "  $_" }
    }
    catch {
        Write-Host "  BitLocker not available on this edition" -ForegroundColor Yellow
    }
    $audit._errors += @{ Section = "BitLocker"; Error = $_.Exception.Message }
}

# =====================================================
# 6. NETWORK CONFIG
# =====================================================
Write-Host ""
Write-Host "=== 6. NETWORK CONFIG ===" -ForegroundColor Cyan

Write-Host "  Running: Get-NetAdapter"
try {
    $adapters = Get-NetAdapter | Select-Object Name, InterfaceDescription, MacAddress, Status, LinkSpeed
    $adapters | Format-Table -AutoSize
    $audit.NetworkAdapters = @($adapters | ForEach-Object {
        [ordered]@{ Name = $_.Name; Description = $_.InterfaceDescription; MAC = $_.MacAddress; Status = "$($_.Status)"; LinkSpeed = $_.LinkSpeed }
    })
    Write-Host "  [OK] Adapters collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "NetworkAdapters"; Error = $_.Exception.Message }
}

Write-Host "  Running: Get-NetIPConfiguration"
try {
    $ipConfigs = Get-NetIPConfiguration | Where-Object { $_.IPv4Address }
    $audit.IPConfig = @($ipConfigs | ForEach-Object {
        [ordered]@{
            Interface = $_.InterfaceAlias
            IPv4 = "$($_.IPv4Address.IPAddress)"
            PrefixLength = $_.IPv4Address.PrefixLength
            Gateway = "$($_.IPv4DefaultGateway.NextHop)"
            DNS = @($_.DNSServer | Where-Object { $_.AddressFamily -eq 2 } | ForEach-Object { $_.ServerAddresses }) | Select-Object -Unique
            DHCPEnabled = (Get-NetIPInterface -InterfaceAlias $_.InterfaceAlias -AddressFamily IPv4 -ErrorAction SilentlyContinue).Dhcp
        }
    })
    $ipConfigs | Format-List InterfaceAlias, IPv4Address, IPv4DefaultGateway, DnsServer
    Write-Host "  [OK] IP config collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "IPConfig"; Error = $_.Exception.Message }
}

# =====================================================
# 7. DOMAIN MEMBERSHIP
# =====================================================
Write-Host ""
Write-Host "=== 7. DOMAIN MEMBERSHIP ===" -ForegroundColor Cyan
Write-Host "  Running: Win32_ComputerSystem + nltest + DirectorySearcher"
try {
    $cs = Get-CimInstance Win32_ComputerSystem
    $domainInfo = [ordered]@{
        DomainJoined = $cs.PartOfDomain
        Domain       = $cs.Domain
        Workgroup    = if (-not $cs.PartOfDomain) { $cs.Workgroup } else { $null }
    }

    if ($cs.PartOfDomain) {
        # Get site
        try {
            $site = nltest /dsgetsite 2>&1 | Select-Object -First 1
            $domainInfo.ADSite = $site.Trim()
        } catch {}

        # Get DC
        try {
            $dcInfo = nltest /dsgetdc:$($cs.Domain) 2>&1
            $dcLine = ($dcInfo | Select-String "DC: \\\\").Line
            if ($dcLine) { $domainInfo.DC = $dcLine.Trim() }
        } catch {}

        # Get computer OU via DirectorySearcher (no RSAT needed)
        try {
            $searcher = New-Object System.DirectoryServices.DirectorySearcher
            $searcher.Filter = "(&(objectClass=computer)(cn=$env:COMPUTERNAME))"
            $result = $searcher.FindOne()
            if ($result) { $domainInfo.ComputerDN = "$($result.Properties['distinguishedname'][0])" }
        } catch {}
    }

    $domainInfo.GetEnumerator() | ForEach-Object { Write-Host "  $($_.Key): $($_.Value)" }
    $audit.DomainMembership = $domainInfo
    Write-Host "  [OK] Domain membership collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "DomainMembership"; Error = $_.Exception.Message }
}

# =====================================================
# 8. CURRENT USER INFO
# =====================================================
Write-Host ""
Write-Host "=== 8. CURRENT USER INFO ===" -ForegroundColor Cyan
Write-Host "  Running: query user, whoami /groups"
try {
    $currentUser = [ordered]@{
        Username = $env:USERNAME
        Domain   = $env:USERDOMAIN
    }

    # Logged-in sessions
    Write-Host "  Active sessions:" -ForegroundColor Yellow
    try {
        $sessions = query user 2>&1
        $sessions | ForEach-Object { Write-Host "    $_" }
    } catch { Write-Host "    Unable to query sessions" }

    # Group memberships (of the account running the script)
    Write-Host "  Group memberships:" -ForegroundColor Yellow
    try {
        $groups = whoami /groups /fo csv 2>&1 | ConvertFrom-Csv -ErrorAction Stop
        $currentUser.Groups = @($groups | ForEach-Object { $_.'Group Name' })
        $groups | ForEach-Object { Write-Host "    $($_.'Group Name')" }
    } catch {
        Write-Host "    Unable to enumerate groups" -ForegroundColor Yellow
    }

    $audit.CurrentUser = $currentUser
    Write-Host "  [OK] User info collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "CurrentUser"; Error = $_.Exception.Message }
}

# =====================================================
# 9. LOCAL USERS
# =====================================================
Write-Host ""
Write-Host "=== 9. LOCAL USERS ===" -ForegroundColor Cyan
Write-Host "  Running: Get-LocalUser"
try {
    $localUsers = Get-LocalUser | Select-Object Name, Enabled, LastLogon, PasswordRequired, PasswordLastSet
    $localUsers | Format-Table -AutoSize
    $audit.LocalUsers = @($localUsers | ForEach-Object {
        [ordered]@{
            Name = $_.Name; Enabled = $_.Enabled
            LastLogon = if ($_.LastLogon) { $_.LastLogon.ToString("yyyy-MM-dd") } else { $null }
            PasswordRequired = $_.PasswordRequired
            PasswordLastSet = if ($_.PasswordLastSet) { $_.PasswordLastSet.ToString("yyyy-MM-dd") } else { $null }
        }
    })
    Write-Host "  [OK] Local users collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message) - trying net user fallback" -ForegroundColor Red
    try { net user } catch {}
    $audit._errors += @{ Section = "LocalUsers"; Error = $_.Exception.Message }
}

# =====================================================
# 10. LOCAL ADMINISTRATORS
# =====================================================
Write-Host ""
Write-Host "=== 10. LOCAL ADMINISTRATORS ===" -ForegroundColor Cyan
Write-Host "  Running: Get-LocalGroupMember Administrators"
try {
    $localAdmins = Get-LocalGroupMember Administrators | Select-Object Name, PrincipalSource, ObjectClass
    $localAdmins | Format-Table -AutoSize
    $audit.LocalAdmins = @($localAdmins | ForEach-Object {
        [ordered]@{ Name = $_.Name; Source = "$($_.PrincipalSource)"; Type = "$($_.ObjectClass)" }
    })
    Write-Host "  [OK] Local admins collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message) - trying net localgroup fallback" -ForegroundColor Red
    try { net localgroup Administrators } catch {}
    $audit._errors += @{ Section = "LocalAdmins"; Error = $_.Exception.Message }
}

# =====================================================
# 11. MAPPED DRIVES
# =====================================================
Write-Host ""
Write-Host "=== 11. MAPPED DRIVES ===" -ForegroundColor Cyan
Write-Host "  Running: Get-PSDrive, net use, HKU registry"
try {
    # Current session drives
    $mappedDrives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.DisplayRoot } |
        Select-Object Name, @{N='UNC';E={$_.DisplayRoot}}
    if ($mappedDrives) {
        Write-Host "  Current session drives:" -ForegroundColor Yellow
        $mappedDrives | Format-Table -AutoSize
    }

    # net use
    Write-Host "  net use output:" -ForegroundColor Yellow
    net use 2>&1 | ForEach-Object { Write-Host "    $_" }

    # Try to get logged-in user's persistent drives via HKU
    $audit.MappedDrives = @()
    try {
        $loggedOnUser = (Get-CimInstance Win32_ComputerSystem).UserName
        if ($loggedOnUser) {
            $sid = (New-Object System.Security.Principal.NTAccount($loggedOnUser)).Translate(
                [System.Security.Principal.SecurityIdentifier]).Value
            try {
                New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS -ErrorAction Stop | Out-Null
                $userDrives = Get-ItemProperty "HKU:\$sid\Network\*" -ErrorAction SilentlyContinue
                if ($userDrives) {
                    Write-Host "  Persistent drives for $loggedOnUser`:" -ForegroundColor Yellow
                    $userDrives | ForEach-Object {
                        Write-Host "    $($_.PSChildName): -> $($_.RemotePath)"
                        $audit.MappedDrives += [ordered]@{ Drive = "$($_.PSChildName):"; UNC = $_.RemotePath; User = $loggedOnUser }
                    }
                }
            }
            finally {
                Remove-PSDrive HKU -ErrorAction SilentlyContinue
            }
        }
    }
    catch {
        Write-Host "  Unable to read user registry for mapped drives: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # Also add from PSDrive
    if ($mappedDrives) {
        foreach ($d in $mappedDrives) {
            if (-not ($audit.MappedDrives | Where-Object { $_.Drive -eq "$($d.Name):" })) {
                $audit.MappedDrives += [ordered]@{ Drive = "$($d.Name):"; UNC = $d.UNC; User = "CurrentSession" }
            }
        }
    }

    Write-Host "  [OK] Mapped drives collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "MappedDrives"; Error = $_.Exception.Message }
}

# =====================================================
# 12. WIFI PROFILES
# =====================================================
Write-Host ""
Write-Host "=== 12. WIFI PROFILES ===" -ForegroundColor Cyan
Write-Host "  Running: netsh wlan show profiles, netsh wlan show interfaces"
try {
    Write-Host "  Saved WiFi profiles:" -ForegroundColor Yellow
    $profiles = netsh wlan show profiles 2>&1
    $profiles | ForEach-Object { Write-Host "    $_" }

    $profileNames = @($profiles | Select-String "All User Profile\s+:\s+(.+)" | ForEach-Object { $_.Matches.Groups[1].Value.Trim() })

    Write-Host ""
    Write-Host "  Current connection:" -ForegroundColor Yellow
    $interfaces = netsh wlan show interfaces 2>&1
    $interfaces | ForEach-Object { Write-Host "    $_" }

    $connectedSSID = ($interfaces | Select-String "SSID\s+:\s+(.+)" | Select-Object -First 1)
    if ($connectedSSID) { $connectedSSID = $connectedSSID.Matches.Groups[1].Value.Trim() }

    $audit.WiFi = [ordered]@{
        SavedProfiles = $profileNames
        ConnectedSSID = $connectedSSID
    }
    Write-Host "  [OK] WiFi profiles collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  (No WiFi adapter or service not running)" -ForegroundColor Yellow
    $audit._errors += @{ Section = "WiFi"; Error = $_.Exception.Message }
}

# =====================================================
# 13. ANTIVIRUS / SECURITY CENTER
# =====================================================
Write-Host ""
Write-Host "=== 13. ANTIVIRUS / SECURITY CENTER ===" -ForegroundColor Cyan

Write-Host "  Running: Get-CimInstance root/SecurityCenter2 AntiVirusProduct"
try {
    $avProducts = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction Stop
    $audit.Antivirus = @($avProducts | ForEach-Object {
        $state = '{0:X6}' -f $_.productState
        $enabled = if ($state.Substring(2,2) -eq '10') { $true } else { $false }
        $upToDate = if ($state.Substring(4,2) -eq '00') { $true } else { $false }

        Write-Host "  $($_.displayName): Enabled=$enabled, UpToDate=$upToDate"

        [ordered]@{
            Name = $_.displayName; Enabled = $enabled; UpToDate = $upToDate
            Path = $_.pathToSignedProductExe
        }
    })
    Write-Host "  [OK] AV products collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "Antivirus"; Error = $_.Exception.Message }
}

Write-Host "  Running: Get-MpComputerStatus (Windows Defender)"
try {
    $defender = Get-MpComputerStatus -ErrorAction Stop
    $defenderInfo = [ordered]@{
        AMRunningMode             = "$($defender.AMRunningMode)"
        AntivirusEnabled          = $defender.AntivirusEnabled
        RealTimeProtection        = $defender.RealTimeProtectionEnabled
        BehaviorMonitor           = $defender.BehaviorMonitorEnabled
        SignatureLastUpdated      = if ($defender.AntivirusSignatureLastUpdated) { $defender.AntivirusSignatureLastUpdated.ToString("yyyy-MM-dd HH:mm") } else { $null }
        LastQuickScan             = if ($defender.QuickScanEndTime) { $defender.QuickScanEndTime.ToString("yyyy-MM-dd HH:mm") } else { $null }
        LastFullScan              = if ($defender.FullScanEndTime) { $defender.FullScanEndTime.ToString("yyyy-MM-dd HH:mm") } else { $null }
    }

    $defenderInfo.GetEnumerator() | ForEach-Object { Write-Host "  Defender $($_.Key): $($_.Value)" }
    $audit.WindowsDefender = $defenderInfo
    Write-Host "  [OK] Defender status collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "WindowsDefender"; Error = $_.Exception.Message }
}

# =====================================================
# 14. PRINTERS
# =====================================================
Write-Host ""
Write-Host "=== 14. PRINTERS ===" -ForegroundColor Cyan
Write-Host "  Running: Get-Printer, Get-PrinterPort"
try {
    $printers = Get-Printer | Select-Object Name, DriverName, PortName, Type, Shared
    $printers | Format-Table -AutoSize

    $ports = Get-PrinterPort | Where-Object { $_.Name -like 'TCP_*' -or $_.Name -like 'IP_*' -or $_.PrinterHostAddress } |
        Select-Object Name, PrinterHostAddress, PortNumber
    if ($ports) {
        Write-Host "  Network printer ports:" -ForegroundColor DarkGray
        $ports | Format-Table -AutoSize
    }

    $audit.Printers = @($printers | ForEach-Object {
        [ordered]@{ Name = $_.Name; Driver = $_.DriverName; Port = $_.PortName; Type = "$($_.Type)"; Shared = $_.Shared }
    })
    $audit.PrinterPorts = @($ports | ForEach-Object {
        [ordered]@{ Name = $_.Name; IP = $_.PrinterHostAddress; Port = $_.PortNumber }
    })
    Write-Host "  [OK] Printers collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "Printers"; Error = $_.Exception.Message }
}

# =====================================================
# 15. INSTALLED SOFTWARE (registry-based)
# =====================================================
Write-Host ""
Write-Host "=== 15. INSTALLED SOFTWARE ===" -ForegroundColor Cyan
Write-Host "  Running: Registry query (HKLM Uninstall keys)"
try {
    $regPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    $software = Get-ItemProperty $regPaths -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName } |
        Select-Object DisplayName, DisplayVersion, Publisher, InstallDate |
        Sort-Object DisplayName -Unique

    $software | Format-Table -AutoSize

    $audit.InstalledSoftware = @($software | ForEach-Object {
        [ordered]@{ Name = $_.DisplayName; Version = $_.DisplayVersion; Publisher = $_.Publisher; InstallDate = $_.InstallDate }
    })
    Write-Host "  [OK] Software collected - $($software.Count) packages" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "InstalledSoftware"; Error = $_.Exception.Message }
}

# =====================================================
# 16. STARTUP PROGRAMS
# =====================================================
Write-Host ""
Write-Host "=== 16. STARTUP PROGRAMS ===" -ForegroundColor Cyan
Write-Host "  Running: Get-CimInstance Win32_StartupCommand"
try {
    $startup = Get-CimInstance Win32_StartupCommand | Select-Object Name, Command, Location, User
    $startup | Format-Table -AutoSize

    $audit.StartupPrograms = @($startup | ForEach-Object {
        [ordered]@{ Name = $_.Name; Command = $_.Command; Location = $_.Location; User = $_.User }
    })
    Write-Host "  [OK] Startup programs collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "StartupPrograms"; Error = $_.Exception.Message }
}

# =====================================================
# 17. WINDOWS UPDATES (last 15)
# =====================================================
Write-Host ""
Write-Host "=== 17. WINDOWS UPDATES (last 15) ===" -ForegroundColor Cyan
Write-Host "  Running: Get-HotFix"
try {
    $updates = Get-HotFix | Sort-Object InstalledOn -Descending -ErrorAction SilentlyContinue | Select-Object -First 15 |
        Select-Object HotFixID, Description, InstalledOn, InstalledBy
    $updates | Format-Table -AutoSize

    $audit.WindowsUpdates = @($updates | ForEach-Object {
        [ordered]@{
            KB = $_.HotFixID; Description = $_.Description
            InstalledOn = if ($_.InstalledOn) { $_.InstalledOn.ToString("yyyy-MM-dd") } else { $null }
        }
    })
    Write-Host "  [OK] Updates collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "WindowsUpdates"; Error = $_.Exception.Message }
}

# =====================================================
# 18. WINDOWS UPDATE SETTINGS
# =====================================================
Write-Host ""
Write-Host "=== 18. WINDOWS UPDATE SETTINGS ===" -ForegroundColor Cyan
Write-Host "  Running: Registry check for WSUS/AU settings"
try {
    $wuSettings = [ordered]@{}

    $wu = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -ErrorAction SilentlyContinue
    if ($wu) {
        $wuSettings.WUServer = $wu.WUServer
        $wuSettings.WUStatusServer = $wu.WUStatusServer
        Write-Host "  WSUS Server: $($wu.WUServer)"
    } else {
        Write-Host "  No WSUS configured (using Windows Update directly)"
        $wuSettings.WUServer = "None (Windows Update)"
    }

    $au = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU' -ErrorAction SilentlyContinue
    if ($au) {
        $auOption = switch ($au.AUOptions) {
            2 {"Notify before download"} 3 {"Auto download, notify install"}
            4 {"Auto download, auto install"} 5 {"Allow local admin to choose"} default {"$($au.AUOptions)"}
        }
        $wuSettings.AutoUpdateOption = $auOption
        Write-Host "  Auto Update: $auOption"
    }

    $audit.WindowsUpdateSettings = $wuSettings
    Write-Host "  [OK] Update settings collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "WindowsUpdateSettings"; Error = $_.Exception.Message }
}

# =====================================================
# 19. NETWORK SHARES
# =====================================================
Write-Host ""
Write-Host "=== 19. NETWORK SHARES (on this machine) ===" -ForegroundColor Cyan
Write-Host "  Running: Get-SmbShare"
try {
    $shares = Get-SmbShare | Where-Object { $_.Name -notlike '*$' }
    if ($shares) {
        $shares | Format-Table Name, Path, Description -AutoSize
        $audit.LocalShares = @($shares | ForEach-Object {
            [ordered]@{ Name = $_.Name; Path = $_.Path; Description = $_.Description }
        })
    } else {
        Write-Host "  No user shares on this machine"
        $audit.LocalShares = @()
    }
    Write-Host "  [OK] Shares collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "LocalShares"; Error = $_.Exception.Message }
}

# =====================================================
# 20. POWER SETTINGS
# =====================================================
Write-Host ""
Write-Host "=== 20. POWER SETTINGS ===" -ForegroundColor Cyan
Write-Host "  Running: powercfg /getactivescheme"
try {
    $powerScheme = powercfg /getactivescheme 2>&1
    Write-Host "  $powerScheme"
    $audit.PowerScheme = "$powerScheme".Trim()
    Write-Host "  [OK] Power settings collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "PowerSettings"; Error = $_.Exception.Message }
}

# =====================================================
# 21. EVENT LOG ERRORS (last 14 days)
# =====================================================
Write-Host ""
Write-Host "=== 21. EVENT LOG ERRORS (last 14 days) ===" -ForegroundColor Cyan
Write-Host "  Running: Get-WinEvent - Critical/Error, last 14 days"
try {
    $errSince = (Get-Date).AddDays(-14)
    $errors = Get-WinEvent -FilterHashtable @{LogName='System'; Level=1,2; StartTime=$errSince} -MaxEvents 30 -ErrorAction Stop |
        Select-Object TimeCreated, Id, LevelDisplayName, ProviderName, @{N='Message';E={$_.Message.Substring(0, [Math]::Min(200, $_.Message.Length))}}
    Write-Host "  $($errors.Count) critical/error events found"
    $errors | Format-Table TimeCreated, Id, ProviderName -AutoSize
    $audit.EventLogErrors = @($errors | ForEach-Object {
        [ordered]@{ Time = $_.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss"); EventId = $_.Id; Level = $_.LevelDisplayName; Source = $_.ProviderName; Message = $_.Message }
    })
    Write-Host "  [OK] Event log errors collected" -ForegroundColor Green
}
catch {
    Write-Host "  No critical/error events in System log (last 14 days)" -ForegroundColor Green
    $audit.EventLogErrors = @()
}

# =====================================================
# 22. EVENT LOG WARNINGS (last 14 days)
# =====================================================
Write-Host ""
Write-Host "=== 22. EVENT LOG WARNINGS (last 14 days) ===" -ForegroundColor Cyan
Write-Host "  Running: Get-WinEvent - Warning level, last 14 days"
try {
    $warnSince = (Get-Date).AddDays(-14)
    $warnings = Get-WinEvent -FilterHashtable @{LogName='System'; Level=3; StartTime=$warnSince} -MaxEvents 50 -ErrorAction Stop |
        Select-Object TimeCreated, Id, ProviderName, @{N='Message';E={$_.Message.Substring(0, [Math]::Min(200, $_.Message.Length))}}
    Write-Host "  $($warnings.Count) warning events found"
    $warnings | Format-Table TimeCreated, Id, ProviderName -AutoSize
    $audit.EventLogWarnings = @($warnings | ForEach-Object {
        [ordered]@{ Time = $_.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss"); EventId = $_.Id; Source = $_.ProviderName; Message = $_.Message }
    })
    Write-Host "  [OK] Event log warnings collected" -ForegroundColor Green
}
catch {
    Write-Host "  No warning events in System log (last 14 days)" -ForegroundColor Green
    $audit.EventLogWarnings = @()
}

# =====================================================
# 23. EVENT LOG SETTINGS
# =====================================================
Write-Host ""
Write-Host "=== 23. EVENT LOG SETTINGS ===" -ForegroundColor Cyan
Write-Host "  Running: Get-WinEvent -ListLog"
try {
    $importantLogs = @('Application', 'Security', 'System', 'Setup', 'Microsoft-Windows-PowerShell/Operational')
    $audit.EventLogSettings = @()
    foreach ($logName in $importantLogs) {
        try {
            $log = Get-WinEvent -ListLog $logName -ErrorAction Stop
            $logInfo = [ordered]@{
                LogName = $log.LogName
                MaxSizeKB = [math]::Round($log.MaximumSizeInBytes / 1KB)
                CurrentSizeKB = [math]::Round($log.FileSize / 1KB)
                RecordCount = $log.RecordCount
                LogMode = "$($log.LogMode)"
                IsEnabled = $log.IsEnabled
            }
            Write-Host "  $($log.LogName): Max=$($logInfo.MaxSizeKB)KB, Mode=$($log.LogMode), Records=$($log.RecordCount)"
            $audit.EventLogSettings += $logInfo
        }
        catch { Write-Host "  $logName`: not available" -ForegroundColor DarkGray }
    }
    Write-Host "  [OK] Event log settings collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "EventLogSettings"; Error = $_.Exception.Message }
}

# =====================================================
# 24. WINDOWS FIREWALL PROFILES
# =====================================================
Write-Host ""
Write-Host "=== 24. WINDOWS FIREWALL PROFILES ===" -ForegroundColor Cyan
Write-Host "  Running: Get-NetFirewallProfile"
try {
    $fwProfiles = Get-NetFirewallProfile -ErrorAction Stop
    $audit.FirewallProfiles = @($fwProfiles | ForEach-Object {
        $p = [ordered]@{
            Profile = "$($_.Name)"; Enabled = $_.Enabled
            DefaultInboundAction = "$($_.DefaultInboundAction)"
            DefaultOutboundAction = "$($_.DefaultOutboundAction)"
        }
        Write-Host "  $($_.Name): Enabled=$($_.Enabled), Inbound=$($_.DefaultInboundAction), Outbound=$($_.DefaultOutboundAction)"
        if (-not $_.Enabled) {
            Write-Host "    [WARN] $($_.Name) firewall profile is DISABLED" -ForegroundColor Red
        }
        $p
    })
    Write-Host "  [OK] Firewall profiles collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "FirewallProfiles"; Error = $_.Exception.Message }
}

# =====================================================
# 25. RDP SECURITY SETTINGS
# =====================================================
Write-Host ""
Write-Host "=== 25. RDP SECURITY SETTINGS ===" -ForegroundColor Cyan
Write-Host "  Running: Registry check for RDP settings"
try {
    $tsReg = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server'
    $rdpReg = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp'

    $rdpEnabled = (Get-ItemProperty $tsReg -ErrorAction Stop).fDenyTSConnections
    $nla = (Get-ItemProperty $rdpReg -ErrorAction SilentlyContinue).UserAuthentication

    $rdpSettings = [ordered]@{
        RDPEnabled = ($rdpEnabled -eq 0)
        NLARequired = ($nla -eq 1)
    }
    $rdpSettings.GetEnumerator() | ForEach-Object { Write-Host "  $($_.Key): $($_.Value)" }

    if ($rdpSettings.RDPEnabled -and -not $rdpSettings.NLARequired) {
        Write-Host "  [WARN] RDP is enabled but NLA is NOT required" -ForegroundColor Yellow
    }

    $audit.RDPSettings = $rdpSettings
    Write-Host "  [OK] RDP settings collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "RDPSettings"; Error = $_.Exception.Message }
}

# =====================================================
# 26. UAC SETTINGS
# =====================================================
Write-Host ""
Write-Host "=== 26. UAC SETTINGS ===" -ForegroundColor Cyan
Write-Host "  Running: Registry check for UAC"
try {
    $uacReg = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -ErrorAction Stop
    $uacSettings = [ordered]@{
        EnableLUA = $uacReg.EnableLUA
        ConsentPromptBehaviorAdmin = switch ($uacReg.ConsentPromptBehaviorAdmin) {
            0 {"Elevate without prompting"} 1 {"Prompt for credentials on secure desktop"}
            2 {"Prompt for consent on secure desktop"} 3 {"Prompt for credentials"}
            4 {"Prompt for consent"} 5 {"Prompt for consent for non-Windows binaries"} default {"Unknown"}
        }
    }
    $uacSettings.GetEnumerator() | ForEach-Object { Write-Host "  $($_.Key): $($_.Value)" }

    if ($uacReg.EnableLUA -ne 1) {
        Write-Host "  [WARN] UAC is DISABLED" -ForegroundColor Red
    }
    $audit.UACSettings = $uacSettings
    Write-Host "  [OK] UAC settings collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "UACSettings"; Error = $_.Exception.Message }
}

# =====================================================
# 27. SCREEN LOCK / INACTIVITY TIMEOUT
# =====================================================
Write-Host ""
Write-Host "=== 27. SCREEN LOCK / INACTIVITY TIMEOUT ===" -ForegroundColor Cyan
Write-Host "  Running: Registry check for screen lock policy"
try {
    $screenLock = [ordered]@{}

    $inactivity = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -ErrorAction SilentlyContinue).InactivityTimeoutSecs
    $screenLock.InactivityTimeoutSecs = $inactivity
    if ($inactivity) {
        Write-Host "  Machine inactivity timeout: $inactivity seconds"
    } else {
        Write-Host "  Machine inactivity timeout: Not configured"
        Write-Host "  [WARN] No machine inactivity timeout - HIPAA requires automatic session lock" -ForegroundColor Yellow
    }

    # Screensaver (may be SYSTEM context - try HKLM policy too)
    $ssPolicy = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Control Panel\Desktop' -ErrorAction SilentlyContinue
    if ($ssPolicy) {
        $screenLock.ScreenSaverGPO = $true
        $screenLock.ScreenSaveTimeout = $ssPolicy.ScreenSaveTimeOut
        $screenLock.ScreenSaverSecure = $ssPolicy.ScreenSaverIsSecure
        Write-Host "  Screensaver GPO: Timeout=$($ssPolicy.ScreenSaveTimeOut)sec, Password=$($ssPolicy.ScreenSaverIsSecure)"
    } else {
        $screenLock.ScreenSaverGPO = $false
        Write-Host "  No screensaver GPO detected"
    }

    $audit.ScreenLockPolicy = $screenLock
    Write-Host "  [OK] Screen lock settings collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "ScreenLock"; Error = $_.Exception.Message }
}

# =====================================================
# 28. USB STORAGE POLICY
# =====================================================
Write-Host ""
Write-Host "=== 28. USB STORAGE POLICY ===" -ForegroundColor Cyan
Write-Host "  Running: Registry check for USB storage restrictions"
try {
    $usbStorage = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Services\USBSTOR' -ErrorAction Stop).Start
    $usbStatus = switch ($usbStorage) { 3 { "Enabled" } 4 { "Disabled" } default { "Unknown ($usbStorage)" } }
    Write-Host "  USB Storage: $usbStatus"

    $usbGPO = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\RemovableStorageDevices' -ErrorAction SilentlyContinue
    if ($usbGPO) { Write-Host "  GPO removable storage restrictions detected" }

    if ($usbStorage -eq 3 -and -not $usbGPO) {
        Write-Host "  [WARN] USB storage is unrestricted - data exfiltration risk" -ForegroundColor Yellow
    }
    $audit.USBStoragePolicy = [ordered]@{ Status = $usbStatus; GPORestrictions = ($null -ne $usbGPO) }
    Write-Host "  [OK] USB storage policy collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "USBStorage"; Error = $_.Exception.Message }
}

# =====================================================
# 29. TLS/SSL CONFIGURATION
# =====================================================
Write-Host ""
Write-Host "=== 29. TLS/SSL CONFIGURATION ===" -ForegroundColor Cyan
Write-Host "  Running: Registry check for TLS protocol versions"
try {
    $tlsVersions = @('SSL 2.0', 'SSL 3.0', 'TLS 1.0', 'TLS 1.1', 'TLS 1.2', 'TLS 1.3')
    $audit.TLSConfig = [ordered]@{}
    foreach ($ver in $tlsVersions) {
        $clientPath = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$ver\Client"
        $enabled = $null
        if (Test-Path $clientPath) {
            $enabled = (Get-ItemProperty $clientPath -ErrorAction SilentlyContinue).Enabled
        }
        $status = if ($null -eq $enabled) { "OS Default" } elseif ($enabled -eq 0) { "Disabled" } else { "Enabled" }
        Write-Host "  $ver`: $status"
        $audit.TLSConfig[$ver] = $status
    }
    Write-Host "  [OK] TLS config collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "TLSConfig"; Error = $_.Exception.Message }
}

# =====================================================
# 30. AUTOPLAY / AUTORUN
# =====================================================
Write-Host ""
Write-Host "=== 30. AUTOPLAY / AUTORUN ===" -ForegroundColor Cyan
Write-Host "  Running: Registry check for AutoPlay/AutoRun"
try {
    $autoplay = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer' -ErrorAction SilentlyContinue).NoDriveTypeAutoRun
    $autorun = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer' -ErrorAction SilentlyContinue).NoAutorun

    if ($autoplay -eq 255) {
        Write-Host "  AutoPlay: Disabled for all drives" -ForegroundColor Green
    } elseif ($autoplay) {
        Write-Host "  AutoPlay: Partially restricted (value: $autoplay)" -ForegroundColor Yellow
    } else {
        Write-Host "  AutoPlay: Not restricted by policy"
        Write-Host "  [WARN] AutoPlay is not disabled - malware risk from USB/optical media" -ForegroundColor Yellow
    }

    $audit.AutoPlay = [ordered]@{ NoDriveTypeAutoRun = $autoplay; NoAutorun = $autorun }
    Write-Host "  [OK] AutoPlay settings collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "AutoPlay"; Error = $_.Exception.Message }
}

# =====================================================
# 31. SERVICES (with expected state check)
# =====================================================
Write-Host ""
Write-Host "=== 31. SERVICES (Auto start - stopped unexpectedly) ===" -ForegroundColor Cyan
Write-Host "  Running: Get-Service - Auto start services that are stopped"
try {
    $stoppedAuto = Get-Service | Where-Object { $_.StartType -eq 'Automatic' -and $_.Status -ne 'Running' } |
        Select-Object Name, DisplayName, Status, StartType
    if ($stoppedAuto) {
        Write-Host "  [WARN] $($stoppedAuto.Count) auto-start services are NOT running:" -ForegroundColor Yellow
        $stoppedAuto | Format-Table -AutoSize
    } else {
        Write-Host "  All auto-start services are running" -ForegroundColor Green
    }
    $audit.StoppedAutoServices = @($stoppedAuto | ForEach-Object {
        [ordered]@{ Name = $_.Name; DisplayName = $_.DisplayName; Status = "$($_.Status)" }
    })
    Write-Host "  [OK] Service state check complete" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "StoppedAutoServices"; Error = $_.Exception.Message }
}

# =====================================================
# 32. FAILED LOGON ATTEMPTS (last 7 days)
# =====================================================
Write-Host ""
Write-Host "=== 32. FAILED LOGON ATTEMPTS (last 7 days) ===" -ForegroundColor Cyan
Write-Host "  Running: Get-WinEvent Security log - Event ID 4625"
$failedSince = (Get-Date).AddDays(-7)
$failedLogons = @(Get-WinEvent -FilterHashtable @{LogName='Security'; Id=4625; StartTime=$failedSince} -MaxEvents 50 -ErrorAction SilentlyContinue)
if ($failedLogons.Count -gt 0) {
    $grouped = $failedLogons | ForEach-Object {
        $xml = [xml]$_.ToXml()
        $targetUser = ($xml.Event.EventData.Data | Where-Object { $_.Name -eq 'TargetUserName' }).'#text'
        $sourceIP = ($xml.Event.EventData.Data | Where-Object { $_.Name -eq 'IpAddress' }).'#text'
        [PSCustomObject]@{ Time = $_.TimeCreated; User = $targetUser; SourceIP = $sourceIP }
    }
    $summary = $grouped | Group-Object User | Sort-Object Count -Descending
    Write-Host "  $($failedLogons.Count) failed logon attempts:" -ForegroundColor Yellow
    $summary | ForEach-Object { Write-Host "    $($_.Name): $($_.Count) failures" -ForegroundColor Yellow }
    $audit.FailedLogons = [ordered]@{
        TotalCount = $failedLogons.Count
        ByUser = @($summary | ForEach-Object { [ordered]@{ User = $_.Name; Count = $_.Count } })
    }
} else {
    Write-Host "  No failed logon attempts" -ForegroundColor Green
    $audit.FailedLogons = [ordered]@{ TotalCount = 0 }
}
Write-Host "  [OK] Failed logons collected" -ForegroundColor Green

# =====================================================
# 33. SECURITY SUMMARY
# =====================================================
Write-Host ""
Write-Host "=== 33. SECURITY SUMMARY ===" -ForegroundColor Cyan
$audit.SecuritySummary = @()

# Check each finding
if ($audit.Activation -and $audit.Activation.StatusCode -ne 1) {
    $audit.SecuritySummary += "OS is NOT fully licensed"
    Write-Host "  [WARN] OS is NOT fully licensed" -ForegroundColor Red
}
if ($audit.FirewallProfiles) {
    $disabledFW = $audit.FirewallProfiles | Where-Object { -not $_.Enabled }
    if ($disabledFW) {
        $audit.SecuritySummary += "Windows Firewall disabled on: $($disabledFW.Profile -join ', ')"
        Write-Host "  [WARN] Windows Firewall disabled on: $($disabledFW.Profile -join ', ')" -ForegroundColor Red
    }
}
if ($audit.BitLocker) {
    $unencrypted = $audit.BitLocker | Where-Object { $_.Protection -eq 'Off' -or $_.VolumeStatus -eq 'FullyDecrypted' }
    if ($unencrypted) {
        $audit.SecuritySummary += "BitLocker not enabled on: $($unencrypted.MountPoint -join ', ')"
        Write-Host "  [WARN] BitLocker not enabled - HIPAA requires encryption at rest" -ForegroundColor Red
    }
}
if (-not $audit.ScreenLockPolicy.InactivityTimeoutSecs -and -not $audit.ScreenLockPolicy.ScreenSaverGPO) {
    $audit.SecuritySummary += "No screen lock / inactivity timeout configured"
    Write-Host "  [WARN] No screen lock configured - HIPAA requires automatic session lock" -ForegroundColor Yellow
}
if ($audit.USBStoragePolicy -and $audit.USBStoragePolicy.Status -eq 'Enabled' -and -not $audit.USBStoragePolicy.GPORestrictions) {
    $audit.SecuritySummary += "USB storage unrestricted"
    Write-Host "  [WARN] USB storage unrestricted - data exfiltration risk" -ForegroundColor Yellow
}
if ($audit.UACSettings -and $audit.UACSettings.EnableLUA -ne 1) {
    $audit.SecuritySummary += "UAC is disabled"
    Write-Host "  [WARN] UAC is disabled" -ForegroundColor Red
}
if ($audit.RDPSettings -and $audit.RDPSettings.RDPEnabled -and -not $audit.RDPSettings.NLARequired) {
    $audit.SecuritySummary += "RDP enabled without NLA"
    Write-Host "  [WARN] RDP enabled without NLA" -ForegroundColor Yellow
}
if ($audit.Antivirus) {
    $noAV = $audit.Antivirus | Where-Object { -not $_.Enabled }
    if ($noAV) {
        # Cross-reference with Get-MpComputerStatus - SecurityCenter2 is unreliable for Defender
        $defenderActive = $audit.WindowsDefender -and $audit.WindowsDefender.AntivirusEnabled -and $audit.WindowsDefender.RealTimeProtection
        $nonDefenderDown = $noAV | Where-Object { $_.Name -ne 'Windows Defender' }
        if ($defenderActive -and -not $nonDefenderDown) {
            # SecurityCenter2 false positive - Defender is actually running
        } else {
            $names = if ($nonDefenderDown) { $nonDefenderDown.Name -join ', ' } else { $noAV.Name -join ', ' }
            $audit.SecuritySummary += "Antivirus not active: $names"
            Write-Host "  [WARN] Antivirus not active" -ForegroundColor Red
        }
    }
}
if ($audit.StoppedAutoServices.Count -gt 0) {
    $audit.SecuritySummary += "$($audit.StoppedAutoServices.Count) auto-start services stopped"
    Write-Host "  [WARN] $($audit.StoppedAutoServices.Count) auto-start services stopped unexpectedly" -ForegroundColor Yellow
}

if ($audit.SecuritySummary.Count -eq 0) {
    Write-Host "  No critical security findings" -ForegroundColor Green
} else {
    Write-Host "  $($audit.SecuritySummary.Count) findings detected - review above" -ForegroundColor Yellow
}

# =====================================================
# 34. BOOT & RECOVERY
# =====================================================
Write-Host ""
Write-Host "=== 34. BOOT & RECOVERY ===" -ForegroundColor Cyan
try {
    $boot = [ordered]@{}

    # Boot mode (UEFI vs Legacy)
    if ($env:firmware_type) {
        $boot.BootMode = $env:firmware_type
    } else {
        try {
            $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
            $boot.BootMode = if ($cs.BootupState) { $cs.BootupState } else { "Unknown" }
        } catch { $boot.BootMode = "Unknown" }
    }

    # Secure Boot (UEFI-only; throws on Legacy BIOS)
    if (Get-Command Confirm-SecureBootUEFI -ErrorAction SilentlyContinue) {
        try { $boot.SecureBootEnabled = [bool](Confirm-SecureBootUEFI -ErrorAction Stop) }
        catch { $boot.SecureBootEnabled = $false; $boot.SecureBootNote = "Not available (Legacy BIOS or unsupported)" }
    } else {
        $boot.SecureBootEnabled = $null
    }

    # TPM
    if (Get-Command Get-Tpm -ErrorAction SilentlyContinue) {
        try {
            $tpm = Get-Tpm -ErrorAction Stop
            $boot.TPM = [ordered]@{
                Present = [bool]$tpm.TpmPresent
                Ready = [bool]$tpm.TpmReady
                Enabled = [bool]$tpm.TpmEnabled
                Activated = [bool]$tpm.TpmActivated
                Owned = [bool]$tpm.TpmOwned
                ManufacturerVersion = "$($tpm.ManufacturerVersion)"
                ManagedAuthLevel = "$($tpm.ManagedAuthLevel)"
            }
        } catch { $boot.TPM = [ordered]@{ Error = $_.Exception.Message } }
    } else {
        try {
            $tpmWmi = Get-CimInstance -Namespace "root\cimv2\security\microsofttpm" -ClassName Win32_Tpm -ErrorAction Stop
            if ($tpmWmi) {
                $boot.TPM = [ordered]@{
                    Present = $true
                    Enabled = [bool]$tpmWmi.IsEnabled_InitialValue
                    Activated = [bool]$tpmWmi.IsActivated_InitialValue
                    Owned = [bool]$tpmWmi.IsOwned_InitialValue
                    SpecVersion = "$($tpmWmi.SpecVersion)"
                    ManufacturerVersion = "$($tpmWmi.ManufacturerVersion)"
                }
            } else { $boot.TPM = [ordered]@{ Present = $false } }
        } catch { $boot.TPM = [ordered]@{ Present = $false; Note = "TPM WMI namespace not available" } }
    }

    # WinRE status
    try {
        $reagent = (& reagentc /info 2>&1) -join "`n"
        $statusMatch = [regex]::Match($reagent, 'Windows RE status:\s+(\w+)')
        $locMatch = [regex]::Match($reagent, 'Windows RE location:\s+(.+)')
        $bcdMatch = [regex]::Match($reagent, 'Boot Configuration Data \(BCD\) identifier:\s+(.+)')
        $boot.WinRE = [ordered]@{
            Status = if ($statusMatch.Success) { $statusMatch.Groups[1].Value.Trim() } else { "Unknown" }
            Location = if ($locMatch.Success) { $locMatch.Groups[1].Value.Trim() } else { "" }
            BCDIdentifier = if ($bcdMatch.Success) { $bcdMatch.Groups[1].Value.Trim() } else { "" }
        }
    } catch { $boot.WinRE = [ordered]@{ Error = $_.Exception.Message } }

    # ESP partition health
    try {
        $espGuid = '{c12a7328-f81f-11d2-ba4b-00a0c93ec93b}'
        $espParts = @(Get-Partition -ErrorAction SilentlyContinue | Where-Object { $_.GptType -eq $espGuid })
        $boot.ESP = @($espParts | ForEach-Object {
            $part = $_
            $vol = $null
            try { $vol = $part | Get-Volume -ErrorAction Stop } catch {}
            [ordered]@{
                DiskNumber = $part.DiskNumber
                PartitionNumber = $part.PartitionNumber
                SizeMB = [math]::Round($part.Size / 1MB, 0)
                FreeMB = if ($vol -and $vol.SizeRemaining) { [math]::Round($vol.SizeRemaining / 1MB, 0) } else { $null }
                FileSystem = if ($vol) { $vol.FileSystem } else { $null }
            }
        })
    } catch { $boot.ESP = @(); $boot.ESPError = $_.Exception.Message }

    # BCD snapshot (parsed)
    try {
        $bcdLines = & bcdedit /enum '{bootmgr}' 2>&1
        $bcdHash = [ordered]@{}
        foreach ($line in $bcdLines) {
            if ($line -match '^([a-zA-Z0-9_]+)\s{2,}(.+)$') {
                $k = $matches[1].Trim()
                $v = $matches[2].Trim()
                if (-not $bcdHash.Contains($k)) { $bcdHash[$k] = $v }
            }
        }
        $boot.BCDBootmgr = $bcdHash
    } catch { $boot.BCDBootmgr = [ordered]@{ Error = $_.Exception.Message } }

    # Recent boot/crash events (last 30 days)
    try {
        $boot30dAgo = (Get-Date).AddDays(-30)
        $bootEvents = @(Get-WinEvent -FilterHashtable @{LogName='System'; Id=41,6008,1074; StartTime=$boot30dAgo} -ErrorAction SilentlyContinue)
        $boot.RecentBootEvents = [ordered]@{
            UnexpectedShutdownsKernel41 = @($bootEvents | Where-Object Id -eq 41).Count
            PreviousShutdownUnexpected6008 = @($bootEvents | Where-Object Id -eq 6008).Count
            PlannedShutdowns1074 = @($bootEvents | Where-Object Id -eq 1074).Count
            Last5 = @($bootEvents | Sort-Object TimeCreated -Descending | Select-Object -First 5 | ForEach-Object {
                [ordered]@{
                    Time = $_.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
                    Id = $_.Id
                    Message = (($_.Message -split "`r?`n")[0]).Trim()
                }
            })
        }
    } catch { $boot.RecentBootEvents = [ordered]@{ Error = $_.Exception.Message } }

    $audit.BootRecovery = $boot

    Write-Host "  Boot mode: $($boot.BootMode), Secure Boot: $($boot.SecureBootEnabled)"
    Write-Host "  TPM present: $($boot.TPM.Present), ready: $($boot.TPM.Ready)"
    Write-Host "  WinRE: $($boot.WinRE.Status)"
    Write-Host "  ESP partitions: $($boot.ESP.Count)"
    Write-Host "  Unexpected shutdowns (30d): $($boot.RecentBootEvents.UnexpectedShutdownsKernel41)"

    if ($boot.SecureBootEnabled -eq $false) { $audit.SecuritySummary += "Secure Boot disabled" }
    if ($boot.TPM.Present -eq $false) { $audit.SecuritySummary += "TPM not present" }
    if ($boot.WinRE.Status -eq "Disabled") { $audit.SecuritySummary += "WinRE disabled (recovery limited)" }
    if ($boot.RecentBootEvents.UnexpectedShutdownsKernel41 -gt 3) {
        $audit.SecuritySummary += "$($boot.RecentBootEvents.UnexpectedShutdownsKernel41) unexpected kernel-power shutdowns in last 30d"
    }
} catch {
    Write-Host "  [ERROR] Boot/Recovery section failed: $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "BootRecovery"; Error = $_.Exception.Message }
}

# =====================================================
# 35. DEFENDER DEEPER
# =====================================================
Write-Host ""
Write-Host "=== 35. DEFENDER DEEPER ===" -ForegroundColor Cyan
try {
    $def = [ordered]@{}

    if (Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue) {
        try {
            $mp = Get-MpComputerStatus -ErrorAction Stop
            $def.Status = [ordered]@{
                AMEngineVersion = "$($mp.AMEngineVersion)"
                AMServiceVersion = "$($mp.AMServiceVersion)"
                AMProductVersion = "$($mp.AMProductVersion)"
                NISEngineVersion = "$($mp.NISEngineVersion)"
                AntispywareSignatureLastUpdated = if ($mp.AntispywareSignatureLastUpdated) { $mp.AntispywareSignatureLastUpdated.ToString("yyyy-MM-dd HH:mm:ss") } else { $null }
                AntivirusSignatureLastUpdated = if ($mp.AntivirusSignatureLastUpdated) { $mp.AntivirusSignatureLastUpdated.ToString("yyyy-MM-dd HH:mm:ss") } else { $null }
                AntivirusSignatureAgeDays = $mp.AntivirusSignatureAge
                NISSignatureAgeDays = $mp.NISSignatureAge
                FullScanAgeDays = $mp.FullScanAge
                QuickScanAgeDays = $mp.QuickScanAge
                FullScanEndTime = if ($mp.FullScanEndTime) { $mp.FullScanEndTime.ToString("yyyy-MM-dd HH:mm:ss") } else { $null }
                QuickScanEndTime = if ($mp.QuickScanEndTime) { $mp.QuickScanEndTime.ToString("yyyy-MM-dd HH:mm:ss") } else { $null }
                IsTamperProtected = if ($mp.PSObject.Properties.Match('IsTamperProtected').Count) { [bool]$mp.IsTamperProtected } else { $null }
                BehaviorMonitorEnabled = [bool]$mp.BehaviorMonitorEnabled
                IoavProtectionEnabled = [bool]$mp.IoavProtectionEnabled
                OnAccessProtectionEnabled = [bool]$mp.OnAccessProtectionEnabled
                AntivirusEnabled = [bool]$mp.AntivirusEnabled
                RealTimeProtectionEnabled = [bool]$mp.RealTimeProtectionEnabled
            }
        } catch { $def.Status = [ordered]@{ Error = $_.Exception.Message } }
    }

    # Tamper Protection (registry fallback)
    try {
        $tpReg = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows Defender\Features" -Name TamperProtection -ErrorAction SilentlyContinue
        if ($tpReg) { $def.TamperProtectionRegValue = $tpReg.TamperProtection }
    } catch {}

    if (Get-Command Get-MpPreference -ErrorAction SilentlyContinue) {
        try {
            $pref = Get-MpPreference -ErrorAction Stop

            # Cloud + sample
            $def.CloudProtection = [ordered]@{
                MAPSReporting = "$($pref.MAPSReporting)"  # 0=Disabled, 1=Basic, 2=Advanced
                SubmitSamplesConsent = "$($pref.SubmitSamplesConsent)"  # 0=AlwaysPrompt, 1=AutoSafe, 2=Never, 3=AutoAll
                CloudBlockLevel = "$($pref.CloudBlockLevel)"
                CloudExtendedTimeout = $pref.CloudExtendedTimeout
            }

            # Controlled Folder Access
            $def.ControlledFolderAccess = [ordered]@{
                EnableControlledFolderAccess = "$($pref.EnableControlledFolderAccess)"  # 0=Disabled, 1=Enabled, 2=AuditMode
                ProtectedFolders = @($pref.ControlledFolderAccessProtectedFolders)
                AllowedApplications = @($pref.ControlledFolderAccessAllowedApplications)
            }

            # ASR rules
            $asrNames = @{
                "56a863a9-875e-4185-98a7-b882c64b5ce5" = "Block abuse of exploited vulnerable signed drivers"
                "7674ba52-37eb-4a4f-a9a1-f0f9a1619a2c" = "Block Adobe Reader from creating child processes"
                "d4f940ab-401b-4efc-aadc-ad5f3c50688a" = "Block all Office applications from creating child processes"
                "9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2" = "Block credential stealing from LSASS"
                "be9ba2d9-53ea-4cdc-84e5-9b1eeee46550" = "Block executable content from email/webmail"
                "01443614-cd74-433a-b99e-2ecdc07bfc25" = "Block executable files unless meeting prevalence/age criteria"
                "5beb7efe-fd9a-4556-801d-275e5ffc04cc" = "Block execution of potentially obfuscated scripts"
                "d3e037e1-3eb8-44c8-a917-57927947596d" = "Block JavaScript/VBScript launching downloaded executable"
                "3b576869-a4ec-4529-8536-b80a7769e899" = "Block Office applications from creating executable content"
                "75668c1f-73b5-4cf0-bb93-3ecf5cb7cc84" = "Block Office applications from injecting code into other processes"
                "26190899-1602-49e8-8b27-eb1d0a1ce869" = "Block Office communication app from creating child processes"
                "e6db77e5-3df2-4cf1-b95a-636979351e5b" = "Block persistence through WMI event subscription"
                "d1e49aac-8f56-4280-b9ba-993a6d77406c" = "Block process creations from PSExec/WMI commands"
                "b2b3f03d-6a65-4f7b-a9c7-1c7ef74a9ba4" = "Block untrusted/unsigned processes that run from USB"
                "92e97fa1-2edf-4476-bdd6-9dd0b4dddc7b" = "Block Win32 API calls from Office macros"
                "c1db55ab-c21a-4637-bb3f-a12568109d35" = "Use advanced ransomware protection"
                "a8f5898e-1dc8-49a9-9878-85004b8a61e6" = "Block Webshell creation for servers"
                "33ddedf1-c6e0-47cb-833e-de6133960387" = "Block rebooting machine in Safe Mode"
                "c0033c00-d16d-4114-a5a0-dc9b3a7d2ceb" = "Block use of copied/impersonated system tools"
            }
            $asrActionNames = @{ 0 = "Disabled"; 1 = "Block"; 2 = "Audit"; 6 = "Warn" }

            $rules = @()
            $ids = @($pref.AttackSurfaceReductionRules_Ids)
            $acts = @($pref.AttackSurfaceReductionRules_Actions)
            for ($i = 0; $i -lt $ids.Count; $i++) {
                $id = "$($ids[$i])"
                $act = if ($i -lt $acts.Count) { [int]$acts[$i] } else { 0 }
                $rules += [ordered]@{
                    Id = $id
                    Name = if ($asrNames.ContainsKey($id)) { $asrNames[$id] } else { "(unknown)" }
                    Action = if ($asrActionNames.ContainsKey($act)) { $asrActionNames[$act] } else { "$act" }
                    ActionCode = $act
                }
            }
            $def.ASRRules = $rules

            # Exclusions (high-value security signal)
            $def.Exclusions = [ordered]@{
                Paths = @($pref.ExclusionPath)
                Extensions = @($pref.ExclusionExtension)
                Processes = @($pref.ExclusionProcess)
                IpAddresses = @($pref.ExclusionIpAddress)
            }
        } catch { $def.PreferenceError = $_.Exception.Message }
    }

    # Threat detection history
    if (Get-Command Get-MpThreatDetection -ErrorAction SilentlyContinue) {
        try {
            $threats = Get-MpThreatDetection -ErrorAction SilentlyContinue
            $def.ThreatHistory = @($threats | Sort-Object InitialDetectionTime -Descending | Select-Object -First 25 | ForEach-Object {
                [ordered]@{
                    ThreatID = "$($_.ThreatID)"
                    ProcessName = "$($_.ProcessName)"
                    Resources = @($_.Resources)
                    InitialDetectionTime = if ($_.InitialDetectionTime) { $_.InitialDetectionTime.ToString("yyyy-MM-dd HH:mm:ss") } else { $null }
                    LastThreatStatusChangeTime = if ($_.LastThreatStatusChangeTime) { $_.LastThreatStatusChangeTime.ToString("yyyy-MM-dd HH:mm:ss") } else { $null }
                    DetectionSourceTypeID = "$($_.DetectionSourceTypeID)"
                    AMProductVersion = "$($_.AMProductVersion)"
                }
            })
            $def.ThreatHistoryCount = @($threats).Count
        } catch { $def.ThreatHistoryError = $_.Exception.Message }
    }

    $audit.DefenderDeeper = $def

    Write-Host "  Sig age (AV): $($def.Status.AntivirusSignatureAgeDays)d, Quick scan age: $($def.Status.QuickScanAgeDays)d, Full scan age: $($def.Status.FullScanAgeDays)d"
    Write-Host "  Tamper protected: $($def.Status.IsTamperProtected) (reg val: $($def.TamperProtectionRegValue))"
    Write-Host "  ASR rules configured: $(@($def.ASRRules).Count) (Block: $(@($def.ASRRules | Where-Object ActionCode -eq 1).Count), Audit: $(@($def.ASRRules | Where-Object ActionCode -eq 2).Count), Disabled: $(@($def.ASRRules | Where-Object ActionCode -eq 0).Count))"
    Write-Host "  Exclusions: paths=$(@($def.Exclusions.Paths).Count) exts=$(@($def.Exclusions.Extensions).Count) procs=$(@($def.Exclusions.Processes).Count)"
    Write-Host "  Threat history entries: $($def.ThreatHistoryCount)"

    # Findings
    if ($def.Status.AntivirusSignatureAgeDays -ne $null -and $def.Status.AntivirusSignatureAgeDays -gt 7) {
        $audit.SecuritySummary += "Defender signatures stale ($($def.Status.AntivirusSignatureAgeDays)d old)"
    }
    if ($def.Status.IsTamperProtected -eq $false) {
        $audit.SecuritySummary += "Defender Tamper Protection disabled"
    }
    if ($def.Status.QuickScanAgeDays -ne $null -and $def.Status.QuickScanAgeDays -gt 14) {
        $audit.SecuritySummary += "Defender quick scan stale ($($def.Status.QuickScanAgeDays)d)"
    }
    if ($def.ControlledFolderAccess.EnableControlledFolderAccess -eq "0") {
        $audit.SecuritySummary += "Controlled Folder Access (anti-ransomware) disabled"
    }
    $blockedAsrCount = @($def.ASRRules | Where-Object ActionCode -eq 1).Count
    if ($blockedAsrCount -lt 5) {
        $audit.SecuritySummary += "Only $blockedAsrCount ASR rules in Block mode (recommend 10+)"
    }
    # Suspicious exclusions
    $suspiciousExcl = @($def.Exclusions.Paths | Where-Object { $_ -match '(?i)\\(temp|appdata|programdata|users\\public)\\' })
    if ($suspiciousExcl.Count -gt 0) {
        $audit.SecuritySummary += "Defender path exclusions in user-writable dirs: $(($suspiciousExcl -join '; '))"
    }
} catch {
    Write-Host "  [ERROR] Defender Deeper section failed: $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "DefenderDeeper"; Error = $_.Exception.Message }
}

# =====================================================
# 36. EDR / SECURITY TOOL OVERLAY
# =====================================================
Write-Host ""
Write-Host "=== 36. EDR / SECURITY TOOL OVERLAY ===" -ForegroundColor Cyan
try {
    $edrCatalog = @(
        @{ Name = "Bitdefender";       Services = @("VSSERV","EPSecurityService","EPProtectedService","EPIntegrationService","EPRedline"); RegPath = "HKLM:\SOFTWARE\Bitdefender" }
        @{ Name = "SentinelOne";       Services = @("SentinelAgent","LogProcessorService","SentinelHelperService"); RegPath = "HKLM:\SOFTWARE\Sentinel Labs" }
        @{ Name = "CrowdStrike Falcon";Services = @("CSFalconService"); RegPath = "HKLM:\SOFTWARE\CrowdStrike" }
        @{ Name = "Webroot";           Services = @("WRSVC","WRCoreService","WRSkyClient"); RegPath = "HKLM:\SOFTWARE\WRData" }
        @{ Name = "Carbon Black";      Services = @("CbDefense","carbonblack","CbProtection","CbComms"); RegPath = "HKLM:\SOFTWARE\CarbonBlack" }
        @{ Name = "Sophos";            Services = @("Sophos Endpoint Defense Service","Sophos Health Service","Sophos MCS Agent","Sophos AutoUpdate Service"); RegPath = "HKLM:\SOFTWARE\Sophos" }
        @{ Name = "ESET";              Services = @("ekrn","EraAgentSvc"); RegPath = "HKLM:\SOFTWARE\ESET" }
        @{ Name = "Malwarebytes";      Services = @("MBAMService","MBAMSvc"); RegPath = "HKLM:\SOFTWARE\Malwarebytes" }
        @{ Name = "Trend Micro";       Services = @("TmCCSF","TmListen","ntrtscan","tmpfw"); RegPath = "HKLM:\SOFTWARE\TrendMicro" }
        @{ Name = "McAfee";            Services = @("masvc","macmnsvc","McAfeeFramework","mfevtps","mfemms"); RegPath = "HKLM:\SOFTWARE\McAfee" }
        @{ Name = "Symantec";          Services = @("ccSvcHst","SmcService","SepMasterService"); RegPath = "HKLM:\SOFTWARE\Symantec" }
        @{ Name = "Huntress";          Services = @("HuntressAgent","HuntressUpdater","HuntressRio"); RegPath = "HKLM:\SOFTWARE\Huntress Labs" }
        @{ Name = "Cylance";           Services = @("CylanceSvc","CylanceUI"); RegPath = "HKLM:\SOFTWARE\Cylance" }
        @{ Name = "ThreatLocker";      Services = @("ThreatLockerService"); RegPath = "HKLM:\SOFTWARE\ThreatLocker" }
        @{ Name = "Defender for Endpoint Sense"; Services = @("Sense"); RegPath = "HKLM:\SOFTWARE\Microsoft\Windows Advanced Threat Protection" }
    )
    $edrFound = @()
    $allServices = @{}
    Get-Service -ErrorAction SilentlyContinue | ForEach-Object { $allServices[$_.Name] = $_ }
    foreach ($prod in $edrCatalog) {
        $matched = @()
        foreach ($svcName in $prod.Services) {
            if ($allServices.ContainsKey($svcName)) {
                $svc = $allServices[$svcName]
                $matched += [ordered]@{ Service = $svcName; Status = "$($svc.Status)"; StartType = "$($svc.StartType)" }
            }
        }
        $regPresent = $false
        try { if (Test-Path $prod.RegPath) { $regPresent = $true } } catch {}
        if ($matched.Count -gt 0 -or $regPresent) {
            $edrFound += [ordered]@{
                Product = $prod.Name
                ServicesPresent = $matched
                RegistryPresent = $regPresent
            }
        }
    }
    $audit.EDRTools = $edrFound
    Write-Host "  EDR/AV products detected: $($edrFound.Count)"
    foreach ($p in $edrFound) {
        $running = @($p.ServicesPresent | Where-Object Status -eq "Running").Count
        Write-Host "    $($p.Product) -- services: $($p.ServicesPresent.Count) ($running running)"
    }
    if ($edrFound.Count -eq 0) {
        Write-Host "  [INFO] No third-party EDR/AV detected (Defender only)" -ForegroundColor Yellow
    }
} catch {
    Write-Host "  [ERROR] EDR overlay section failed: $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "EDRTools"; Error = $_.Exception.Message }
}

# =====================================================
# 37. SCHEDULED TASKS (with suspicious flags)
# =====================================================
Write-Host ""
Write-Host "=== 37. SCHEDULED TASKS ===" -ForegroundColor Cyan
try {
    $tasks = @()
    $suspiciousTasks = @()
    $created30dAgo = (Get-Date).AddDays(-30)
    $allTasks = @(Get-ScheduledTask -ErrorAction SilentlyContinue)
    foreach ($t in $allTasks) {
        $info = $null
        try { $info = $t | Get-ScheduledTaskInfo -ErrorAction Stop } catch {}
        $actions = @($t.Actions | ForEach-Object {
            [ordered]@{
                Type = $_.GetType().Name
                Execute = "$($_.Execute)"
                Arguments = "$($_.Arguments)"
                WorkingDirectory = "$($_.WorkingDirectory)"
            }
        })
        # Suspicious patterns
        $isSuspicious = $false
        $suspReasons = @()
        foreach ($a in $actions) {
            $cmd = "$($a.Execute) $($a.Arguments)"
            if ($cmd -match '(?i)\\(temp|appdata\\local\\temp|programdata\\temp|users\\public)\\') { $isSuspicious=$true; $suspReasons += "Runs from user-writable temp" }
            if ($cmd -match '(?i)powershell.*\s+-(e|en|enc|encod|encode|encodedcommand)\b') { $isSuspicious=$true; $suspReasons += "PowerShell -EncodedCommand" }
            if ($cmd -match '(?i)\bmshta\.exe') { $isSuspicious=$true; $suspReasons += "mshta.exe (HTA execution)" }
            if ($cmd -match '(?i)\brundll32\.exe.*,([A-Z][a-z]+)') { $suspReasons += "rundll32 with custom ordinal" }
            if ($cmd -match '(?i)\bcertutil\.exe.*-(decode|urlcache|f)\b') { $isSuspicious=$true; $suspReasons += "certutil decode/download" }
            if ($cmd -match '(?i)\bbitsadmin\.exe.*\/transfer') { $isSuspicious=$true; $suspReasons += "bitsadmin transfer" }
            if ($cmd -match '(?i)\b(curl|wget|iwr|invoke-webrequest|invoke-restmethod)\b.*\bhttps?://') { $isSuspicious=$true; $suspReasons += "Inline HTTP download" }
        }
        $authorOk = $true
        try {
            if ($t.Author -and $t.Author -notmatch '^(Microsoft|Windows|\\?\\?\\?|NT AUTHORITY|SYSTEM|S-1-5-)') {
                $authorOk = $true  # custom author -- may be normal
            }
        } catch {}
        $createdRecent = $false
        if ($info -and $info.LastRunTime -and ($info.LastRunTime -gt $created30dAgo)) { $createdRecent = $true }

        $task = [ordered]@{
            TaskName = $t.TaskName
            TaskPath = $t.TaskPath
            State = "$($t.State)"
            Author = $t.Author
            Description = $t.Description
            Actions = $actions
            LastRunTime = if ($info -and $info.LastRunTime) { $info.LastRunTime.ToString("yyyy-MM-dd HH:mm:ss") } else { $null }
            NextRunTime = if ($info -and $info.NextRunTime) { $info.NextRunTime.ToString("yyyy-MM-dd HH:mm:ss") } else { $null }
            LastTaskResult = if ($info) { "$($info.LastTaskResult)" } else { $null }
        }
        if ($isSuspicious) {
            $task.SuspiciousReasons = $suspReasons
            $suspiciousTasks += $task
        }
        # Keep all non-Microsoft tasks; Microsoft tasks only if suspicious
        if ($t.TaskPath -notmatch '^\\Microsoft\\' -or $isSuspicious) {
            $tasks += $task
        }
    }
    $audit.ScheduledTasks = [ordered]@{
        TotalCount = $allTasks.Count
        ReturnedCount = $tasks.Count
        SuspiciousCount = $suspiciousTasks.Count
        SuspiciousTasks = $suspiciousTasks
        AllNonMicrosoft = $tasks
    }
    Write-Host "  Total tasks: $($allTasks.Count), Non-Microsoft: $($tasks.Count), Suspicious: $($suspiciousTasks.Count)"
    if ($suspiciousTasks.Count -gt 0) {
        $audit.SecuritySummary += "$($suspiciousTasks.Count) suspicious scheduled task(s) flagged"
        foreach ($st in $suspiciousTasks) {
            Write-Host "    [SUSP] $($st.TaskPath)$($st.TaskName) -- $($st.SuspiciousReasons -join '; ')" -ForegroundColor Red
        }
    }
} catch {
    Write-Host "  [ERROR] Scheduled tasks section failed: $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "ScheduledTasks"; Error = $_.Exception.Message }
}

# =====================================================
# 38. SERVICES FULL INVENTORY (with suspicious flags)
# =====================================================
Write-Host ""
Write-Host "=== 38. SERVICES INVENTORY ===" -ForegroundColor Cyan
try {
    $allSvc = @(Get-CimInstance -ClassName Win32_Service -ErrorAction SilentlyContinue)
    $suspSvc = @()
    foreach ($s in $allSvc) {
        $reasons = @()
        $path = "$($s.PathName)"
        # Extract just the executable from the PathName
        $exe = ""
        if ($path -match '^"([^"]+)"') { $exe = $matches[1] }
        elseif ($path -match '^(\S+\.exe)') { $exe = $matches[1] }
        else { $exe = ($path -split '\s')[0] }

        # Path with spaces but not quoted (privilege-escalation classic)
        if ($path -and $path -notmatch '^"' -and $path -match '\s' -and $path -match '^([A-Z]:\\[^"]+\s+[^"]+\.exe)') {
            $reasons += "Unquoted path with spaces"
        }
        # Binary in user-writable dir
        if ($exe -match '(?i)\\(users|programdata|temp|public|appdata)\\') {
            $reasons += "Binary in user-writable directory: $exe"
        }
        # Not in standard system paths
        if ($exe -and $exe -notmatch '(?i)^[A-Z]:\\(Windows|Program Files|Program Files \(x86\))') {
            if ($exe -notmatch '(?i)^[A-Z]:\\Users') {
                $reasons += "Binary outside standard system paths: $exe"
            }
        }
        # StartName not standard
        $startName = "$($s.StartName)"
        if ($startName -and $startName -notmatch '^(LocalSystem|NT AUTHORITY|NT Service|.*\\LocalService|.*\\NetworkService)$' -and $startName -ne '') {
            # Custom user account -- worth noting
        }

        if ($reasons.Count -gt 0) {
            $suspSvc += [ordered]@{
                Name = $s.Name
                DisplayName = $s.DisplayName
                State = "$($s.State)"
                StartMode = "$($s.StartMode)"
                StartName = $startName
                PathName = $path
                ProcessId = $s.ProcessId
                SuspiciousReasons = $reasons
            }
        }
    }
    $audit.ServicesInventory = [ordered]@{
        TotalServices = $allSvc.Count
        SuspiciousCount = $suspSvc.Count
        SuspiciousServices = $suspSvc
    }
    Write-Host "  Total services: $($allSvc.Count), Suspicious: $($suspSvc.Count)"
    if ($suspSvc.Count -gt 0) {
        $audit.SecuritySummary += "$($suspSvc.Count) suspicious service(s) flagged (path/binary anomalies)"
        foreach ($s in $suspSvc) {
            Write-Host "    [SUSP] $($s.Name) ($($s.DisplayName)) -- $($s.SuspiciousReasons -join '; ')" -ForegroundColor Red
        }
    }
} catch {
    Write-Host "  [ERROR] Services inventory section failed: $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "ServicesInventory"; Error = $_.Exception.Message }
}

# =====================================================
# 39. PERSISTENCE - REGISTRY RUN/IFEO/WINLOGON + WMI SUBSCRIPTIONS
# =====================================================
Write-Host ""
Write-Host "=== 39. REGISTRY PERSISTENCE + WMI SUBSCRIPTIONS ===" -ForegroundColor Cyan
try {
    $persistence = [ordered]@{}

    $runKeyPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run",
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Run",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\RunOnce",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce"
    )
    $runEntries = @()
    foreach ($p in $runKeyPaths) {
        try {
            if (Test-Path $p) {
                $vals = Get-ItemProperty -Path $p -ErrorAction SilentlyContinue
                if ($vals) {
                    $vals.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' } | ForEach-Object {
                        $runEntries += [ordered]@{
                            Hive = $p
                            ValueName = $_.Name
                            Command = "$($_.Value)"
                        }
                    }
                }
            }
        } catch {}
    }
    $persistence.RunKeys = $runEntries

    # Winlogon Userinit + Shell
    try {
        $wl = Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Winlogon" -ErrorAction SilentlyContinue
        $persistence.Winlogon = [ordered]@{
            Userinit = "$($wl.Userinit)"
            Shell = "$($wl.Shell)"
        }
        if ($wl.Userinit -and $wl.Userinit -notmatch '(?i)^[A-Z]:\\Windows\\system32\\userinit\.exe,?\s*$') {
            $audit.SecuritySummary += "Winlogon Userinit modified: $($wl.Userinit)"
        }
        if ($wl.Shell -and $wl.Shell -notmatch '(?i)^explorer\.exe$') {
            $audit.SecuritySummary += "Winlogon Shell modified: $($wl.Shell)"
        }
    } catch { $persistence.Winlogon = [ordered]@{ Error = $_.Exception.Message } }

    # Image File Execution Options - look for Debugger value (debugger hijack)
    try {
        $ifeoBase = "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"
        $ifeoDebuggers = @()
        if (Test-Path $ifeoBase) {
            Get-ChildItem -Path $ifeoBase -ErrorAction SilentlyContinue | ForEach-Object {
                $sub = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
                if ($sub.Debugger) {
                    $ifeoDebuggers += [ordered]@{
                        TargetExecutable = $_.PSChildName
                        Debugger = "$($sub.Debugger)"
                    }
                }
            }
        }
        $persistence.IFEODebuggers = $ifeoDebuggers
        if ($ifeoDebuggers.Count -gt 0) {
            $audit.SecuritySummary += "IFEO Debugger hijack(s) found: $($ifeoDebuggers.Count)"
        }
    } catch { $persistence.IFEODebuggers = @() }

    # WMI subscriptions (classic APT persistence)
    $wmiSubs = [ordered]@{}
    try {
        $filters = @(Get-CimInstance -Namespace root\subscription -ClassName __EventFilter -ErrorAction SilentlyContinue)
        $consumers = @(Get-CimInstance -Namespace root\subscription -ClassName __EventConsumer -ErrorAction SilentlyContinue)
        $bindings = @(Get-CimInstance -Namespace root\subscription -ClassName __FilterToConsumerBinding -ErrorAction SilentlyContinue)
        $wmiSubs.EventFilters = @($filters | ForEach-Object { [ordered]@{ Name = $_.Name; Query = "$($_.Query)"; QueryLanguage = "$($_.QueryLanguage)"; EventNamespace = "$($_.EventNamespace)" } })
        $wmiSubs.EventConsumers = @($consumers | ForEach-Object {
            [ordered]@{
                Name = $_.Name
                Class = $_.CimClass.CimClassName
                CommandLineTemplate = "$($_.CommandLineTemplate)"
                ScriptText = "$($_.ScriptText)"
                ExecutablePath = "$($_.ExecutablePath)"
            }
        })
        $wmiSubs.Bindings = @($bindings | ForEach-Object { [ordered]@{ Filter = "$($_.Filter)"; Consumer = "$($_.Consumer)" } })
        # Anything user-defined here is suspicious
        $userFilters = @($filters | Where-Object { $_.Name -notmatch '^(BVTFilter|SCM Event Log Filter|NTEventLogProvider)$' })
        if ($userFilters.Count -gt 0) {
            $audit.SecuritySummary += "$($userFilters.Count) user-defined WMI event filter(s) -- review for persistence"
        }
    } catch { $wmiSubs.Error = $_.Exception.Message }
    $persistence.WMISubscriptions = $wmiSubs

    $audit.RegistryPersistence = $persistence
    Write-Host "  Run-key entries: $($runEntries.Count)"
    Write-Host "  IFEO debugger hijacks: $($persistence.IFEODebuggers.Count)"
    Write-Host "  WMI EventFilters: $(@($wmiSubs.EventFilters).Count), Consumers: $(@($wmiSubs.EventConsumers).Count), Bindings: $(@($wmiSubs.Bindings).Count)"
} catch {
    Write-Host "  [ERROR] Registry/WMI persistence section failed: $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "RegistryPersistence"; Error = $_.Exception.Message }
}

# =====================================================
# 40. RECENTLY MODIFIED FILES (suspicious paths, last 7d)
# =====================================================
Write-Host ""
Write-Host "=== 40. RECENTLY MODIFIED FILES ===" -ForegroundColor Cyan
try {
    $scanDirs = @()
    foreach ($d in @($env:TEMP, "$env:LOCALAPPDATA\Temp", $env:APPDATA, $env:LOCALAPPDATA, $env:PROGRAMDATA, $env:PUBLIC, "$env:SystemRoot\Temp")) {
        if ($d -and (Test-Path $d)) { $scanDirs += $d }
    }
    $scanDirs = $scanDirs | Select-Object -Unique
    $sevenDaysAgo = (Get-Date).AddDays(-7)
    $execExtensions = '\.(exe|dll|ps1|vbs|vbe|js|jse|hta|jar|bat|cmd|scr|msi|cpl|wsh|wsf|lnk|pif)$'
    $excludePathPatterns = @(
        '(?i)\\(BraveSoftware|Google\\Chrome|Microsoft\\Edge|Mozilla\\Firefox|Microsoft\\Teams|GitHub Desktop|Slack|Discord|Spotify|Zoom)\\',
        '(?i)\\(packages|cache|crashpad|GPUCache|ShaderCache|Code Cache|Cache_Data|webrtc|service_worker|IndexedDB|Local Storage)\\',
        '(?i)\\(Microsoft\\Windows\\(WebCache|INetCache|Network|Notifications|Cookies|Explorer\\thumbcache))\\',
        '(?i)\\(NuGet|pip|npm-cache|yarn-cache|.gradle|.m2|.cargo|.rustup)\\'
    )
    $allFiles = @()
    foreach ($d in $scanDirs) {
        try {
            $found = @(Get-ChildItem -Path $d -File -Recurse -Force -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -gt $sevenDaysAgo })
            $allFiles += $found
        } catch {}
    }
    # De-noise
    $filtered = $allFiles | Where-Object {
        $f = $_.FullName
        $exclude = $false
        foreach ($pat in $excludePathPatterns) { if ($f -match $pat) { $exclude = $true; break } }
        -not $exclude
    }
    $top50 = $filtered | Sort-Object LastWriteTime -Descending | Select-Object -First 50
    $execFiles = @($top50 | Where-Object { $_.Name -match $execExtensions })
    $audit.RecentlyModifiedFiles = [ordered]@{
        ScanDirectories = $scanDirs
        TotalScanned = $allFiles.Count
        AfterFilter = $filtered.Count
        Top50 = @($top50 | ForEach-Object {
            [ordered]@{
                Path = $_.FullName
                SizeBytes = $_.Length
                LastWriteTime = $_.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                Extension = $_.Extension
            }
        })
        ExecutableInUserDirs = @($execFiles | ForEach-Object {
            [ordered]@{
                Path = $_.FullName
                SizeBytes = $_.Length
                LastWriteTime = $_.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                Extension = $_.Extension
            }
        })
    }
    Write-Host "  Scanned: $($allFiles.Count) files (after filter: $($filtered.Count)) across $($scanDirs.Count) dirs"
    Write-Host "  Executable-extension files in user-writable dirs (7d): $($execFiles.Count)"
    if ($execFiles.Count -gt 0) {
        $audit.SecuritySummary += "$($execFiles.Count) executable file(s) recently modified in user-writable dirs"
        foreach ($f in ($execFiles | Select-Object -First 5)) {
            Write-Host "    [SUSP] $($f.FullName) -- $($f.LastWriteTime)" -ForegroundColor Yellow
        }
    }
} catch {
    Write-Host "  [ERROR] Recently modified files section failed: $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "RecentlyModifiedFiles"; Error = $_.Exception.Message }
}

# =====================================================
# 41. REMOTE ACCESS TOOLS INVENTORY
# =====================================================
Write-Host ""
Write-Host "=== 41. REMOTE ACCESS TOOLS ===" -ForegroundColor Cyan
try {
    $ratCatalog = @(
        @{ Name = "ScreenConnect/ConnectWise Control"; ServicePattern = '^ScreenConnect Client'; Paths = @("C:\Program Files (x86)\ScreenConnect Client*", "C:\Program Files\ScreenConnect Client*") }
        @{ Name = "TeamViewer";    ServicePattern = '^TeamViewer';     Paths = @("C:\Program Files\TeamViewer", "C:\Program Files (x86)\TeamViewer") }
        @{ Name = "AnyDesk";       ServicePattern = '^AnyDesk';        Paths = @("C:\Program Files\AnyDesk", "C:\Program Files (x86)\AnyDesk", "C:\ProgramData\AnyDesk") }
        @{ Name = "Splashtop";     ServicePattern = '^Splashtop';      Paths = @("C:\Program Files (x86)\Splashtop", "C:\Program Files\Splashtop") }
        @{ Name = "LogMeIn";       ServicePattern = '^(LMI|LogMeIn)';  Paths = @("C:\Program Files (x86)\LogMeIn", "C:\Program Files\LogMeIn") }
        @{ Name = "GoToAssist/GoToMyPC"; ServicePattern = '^(g2m|GoTo)'; Paths = @("C:\Program Files (x86)\Citrix\GoToAssist Expert*", "C:\Program Files (x86)\GoToMyPC") }
        @{ Name = "Supremo";       ServicePattern = '^SupremoService'; Paths = @("C:\Program Files (x86)\SupremoRemoteDesktop", "C:\Program Files\SupremoRemoteDesktop") }
        @{ Name = "RustDesk";      ServicePattern = '^RustDesk';       Paths = @("C:\Program Files\RustDesk") }
        @{ Name = "Atera";         ServicePattern = '^AteraAgent';     Paths = @("C:\Program Files\ATERA Networks") }
        @{ Name = "NinjaOne/NinjaRMM"; ServicePattern = '^NinjaRMMAgent'; Paths = @("C:\Program Files (x86)\NinjaRMMAgent", "C:\Program Files\NinjaRMMAgent") }
        @{ Name = "Action1";       ServicePattern = '^Action1';        Paths = @("C:\Program Files (x86)\Action1") }
        @{ Name = "Tailscale";     ServicePattern = '^Tailscale';      Paths = @("C:\Program Files\Tailscale") }
        @{ Name = "ZeroTier";      ServicePattern = '^ZeroTierOneService'; Paths = @("C:\ProgramData\ZeroTier") }
        @{ Name = "RemotePC";      ServicePattern = '^RemotePC';       Paths = @("C:\Program Files\RemotePC") }
        @{ Name = "Kaseya VSA";    ServicePattern = '^KaseyaAgent';    Paths = @("C:\Program Files (x86)\Kaseya") }
        @{ Name = "Datto RMM";     ServicePattern = '^CagService';     Paths = @("C:\Program Files (x86)\CentraStage") }
        @{ Name = "N-able N-central"; ServicePattern = '^Windows Agent'; Paths = @("C:\Program Files (x86)\N-able Technologies") }
        @{ Name = "Chrome Remote Desktop"; ServicePattern = '^chromoting'; Paths = @("C:\Program Files (x86)\Google\Chrome Remote Desktop") }
        @{ Name = "GuruRMM";       ServicePattern = '^GuruRMM';        Paths = @("C:\Program Files\GuruRMM", "C:\ProgramData\GuruRMM") }
    )
    $allSvcRat = @{}
    Get-Service -ErrorAction SilentlyContinue | ForEach-Object { $allSvcRat[$_.Name] = $_ }
    $rats = @()
    foreach ($r in $ratCatalog) {
        $matchedSvcs = @($allSvcRat.Values | Where-Object { $_.Name -match $r.ServicePattern -or $_.DisplayName -match $r.ServicePattern })
        $matchedPaths = @()
        foreach ($p in $r.Paths) {
            try { if (Get-ChildItem -Path $p -ErrorAction SilentlyContinue) { $matchedPaths += $p } } catch {}
        }
        if ($matchedSvcs.Count -gt 0 -or $matchedPaths.Count -gt 0) {
            $rats += [ordered]@{
                Product = $r.Name
                Services = @($matchedSvcs | ForEach-Object { [ordered]@{ Name = $_.Name; DisplayName = $_.DisplayName; Status = "$($_.Status)" } })
                InstallPaths = $matchedPaths
            }
        }
    }
    $audit.RemoteAccessTools = $rats
    Write-Host "  Remote-access tools detected: $($rats.Count)"
    foreach ($rat in $rats) {
        $running = @($rat.Services | Where-Object Status -eq "Running").Count
        Write-Host "    $($rat.Product) -- services: $($rat.Services.Count) ($running running) -- paths: $($rat.InstallPaths.Count)"
    }
    if ($rats.Count -gt 4) {
        $audit.SecuritySummary += "$($rats.Count) remote-access tools installed -- review whether all are sanctioned"
    }
} catch {
    Write-Host "  [ERROR] Remote access tools section failed: $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "RemoteAccessTools"; Error = $_.Exception.Message }
}

# =====================================================
# 42. NETWORK DEEPER (listening ports, established outbound, hosts, proxy, DNS)
# =====================================================
Write-Host ""
Write-Host "=== 42. NETWORK DEEPER ===" -ForegroundColor Cyan
try {
    $net = [ordered]@{}

    # Listening TCP with owning process
    try {
        $procIndex = @{}
        Get-Process -ErrorAction SilentlyContinue | ForEach-Object { $procIndex[$_.Id] = $_ }
        $listening = @(Get-NetTCPConnection -State Listen -ErrorAction SilentlyContinue | ForEach-Object {
            $proc = $procIndex[[int]$_.OwningProcess]
            [ordered]@{
                LocalAddress = $_.LocalAddress
                LocalPort = $_.LocalPort
                OwningProcessId = $_.OwningProcess
                ProcessName = if ($proc) { $proc.ProcessName } else { "" }
                ProcessPath = if ($proc -and $proc.Path) { $proc.Path } else { "" }
            }
        } | Sort-Object LocalPort)
        $net.ListeningTCP = $listening
    } catch { $net.ListeningTCPError = $_.Exception.Message; $net.ListeningTCP = @() }

    # Established outbound to public IPs
    try {
        $procIndex2 = @{}
        Get-Process -ErrorAction SilentlyContinue | ForEach-Object { $procIndex2[$_.Id] = $_ }
        $established = @(Get-NetTCPConnection -State Established -ErrorAction SilentlyContinue | Where-Object {
            $r = $_.RemoteAddress
            -not ($r -match '^(127\.|10\.|192\.168\.|169\.254\.|::1$|fe80:|0\.0\.0\.0$)' -or
                  $r -match '^172\.(1[6-9]|2[0-9]|3[01])\.')
        } | ForEach-Object {
            $proc = $procIndex2[[int]$_.OwningProcess]
            [ordered]@{
                LocalPort = $_.LocalPort
                RemoteAddress = $_.RemoteAddress
                RemotePort = $_.RemotePort
                ProcessName = if ($proc) { $proc.ProcessName } else { "" }
                ProcessPath = if ($proc -and $proc.Path) { $proc.Path } else { "" }
            }
        } | Sort-Object ProcessName, RemoteAddress | Select-Object -First 100)
        $net.EstablishedOutbound = $established
    } catch { $net.EstablishedOutboundError = $_.Exception.Message; $net.EstablishedOutbound = @() }

    # HOSTS file
    try {
        $hostsPath = "$env:windir\System32\drivers\etc\hosts"
        $hostsContent = Get-Content $hostsPath -ErrorAction SilentlyContinue
        $hostsActive = @($hostsContent | Where-Object { $_ -and ($_ -notmatch '^\s*#') -and ($_.Trim() -ne '') })
        $net.HostsFile = [ordered]@{
            ActiveEntryCount = $hostsActive.Count
            ActiveEntries = $hostsActive
        }
        if ($hostsActive.Count -gt 0) {
            $audit.SecuritySummary += "HOSTS file has $($hostsActive.Count) non-default active entries"
        }
    } catch { $net.HostsFile = [ordered]@{ Error = $_.Exception.Message } }

    # Proxy (system + per-user)
    try {
        $proxyHKCU = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings" -ErrorAction SilentlyContinue
        $net.ProxyHKCU = [ordered]@{
            ProxyEnable = $proxyHKCU.ProxyEnable
            ProxyServer = "$($proxyHKCU.ProxyServer)"
            AutoConfigURL = "$($proxyHKCU.AutoConfigURL)"
            ProxyOverride = "$($proxyHKCU.ProxyOverride)"
        }
        $winhttp = (& netsh winhttp show proxy 2>&1) -join "`n"
        $net.WinHTTPProxy = $winhttp.Trim()
        if ($proxyHKCU.ProxyEnable -eq 1 -and $proxyHKCU.ProxyServer) {
            $audit.SecuritySummary += "User proxy configured: $($proxyHKCU.ProxyServer) -- verify it's expected"
        }
        if ($proxyHKCU.AutoConfigURL) {
            $audit.SecuritySummary += "Proxy auto-config URL set: $($proxyHKCU.AutoConfigURL)"
        }
    } catch { $net.ProxyError = $_.Exception.Message }

    # DNS servers per active interface
    try {
        $dns = @(Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.ServerAddresses.Count -gt 0 } | ForEach-Object {
            [ordered]@{
                InterfaceAlias = $_.InterfaceAlias
                InterfaceIndex = $_.InterfaceIndex
                Servers = @($_.ServerAddresses)
            }
        })
        $net.DNSServers = $dns
    } catch { $net.DNSServers = @() }

    # Network connection profiles per interface
    try {
        $net.ConnectionProfiles = @(Get-NetConnectionProfile -ErrorAction SilentlyContinue | ForEach-Object {
            [ordered]@{
                InterfaceAlias = $_.InterfaceAlias
                NetworkCategory = "$($_.NetworkCategory)"
                IPv4Connectivity = "$($_.IPv4Connectivity)"
                Name = $_.Name
            }
        })
        # Flag domain-joined machine on Public network
        $publicProfiles = @($net.ConnectionProfiles | Where-Object NetworkCategory -eq "Public")
        if ($audit.DomainMembership -and $audit.DomainMembership.PartOfDomain -and $publicProfiles.Count -gt 0) {
            $audit.SecuritySummary += "Domain-joined machine has interface(s) on Public network profile"
        }
    } catch { $net.ConnectionProfiles = @() }

    $audit.NetworkDeeper = $net
    Write-Host "  Listening TCP: $(@($net.ListeningTCP).Count), Outbound (public): $(@($net.EstablishedOutbound).Count)"
    Write-Host "  HOSTS active entries: $($net.HostsFile.ActiveEntryCount)"
    Write-Host "  HKCU proxy enabled: $($net.ProxyHKCU.ProxyEnable)"
    Write-Host "  DNS interfaces with servers: $(@($net.DNSServers).Count)"
} catch {
    Write-Host "  [ERROR] Network Deeper section failed: $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "NetworkDeeper"; Error = $_.Exception.Message }
}

# =====================================================
# 43. BROWSER HYGIENE (Edge, Chrome, Brave, Firefox)
# =====================================================
Write-Host ""
Write-Host "=== 43. BROWSER HYGIENE ===" -ForegroundColor Cyan
try {
    $browsers = @()

    function Get-ChromiumExtensions {
        param($ProfileDir, $BrowserName)
        $extDir = Join-Path $ProfileDir "Extensions"
        if (-not (Test-Path $extDir)) { return @() }
        $exts = @()
        Get-ChildItem -Path $extDir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $extId = $_.Name
            $verDir = Get-ChildItem -Path $_.FullName -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 1
            if ($verDir) {
                $manifestPath = Join-Path $verDir.FullName "manifest.json"
                if (Test-Path $manifestPath) {
                    try {
                        $manifest = Get-Content $manifestPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                        $name = "$($manifest.name)"
                        # Resolve __MSG_ name from default_locale messages.json if needed
                        if ($name -match '^__MSG_(.+)__$') {
                            $msgKey = $matches[1]
                            $defLocale = if ($manifest.default_locale) { $manifest.default_locale } else { "en" }
                            $msgPath = Join-Path $verDir.FullName "_locales\$defLocale\messages.json"
                            if (Test-Path $msgPath) {
                                try {
                                    $messages = Get-Content $msgPath -Raw | ConvertFrom-Json
                                    $msgEntry = $messages.PSObject.Properties | Where-Object { $_.Name -ieq $msgKey } | Select-Object -First 1
                                    if ($msgEntry) { $name = "$($msgEntry.Value.message)" }
                                } catch {}
                            }
                        }
                        $perms = @()
                        if ($manifest.permissions) { $perms += @($manifest.permissions) }
                        if ($manifest.host_permissions) { $perms += @($manifest.host_permissions) }
                        $exts += [ordered]@{
                            Browser = $BrowserName
                            Id = $extId
                            Name = $name
                            Version = "$($manifest.version)"
                            Permissions = $perms
                            UpdateURL = "$($manifest.update_url)"
                        }
                    } catch {
                        $exts += [ordered]@{ Browser = $BrowserName; Id = $extId; Name = "(manifest unreadable)"; Error = $_.Exception.Message }
                    }
                }
            }
        }
        return $exts
    }

    # Edge
    $edgePath = "$env:LOCALAPPDATA\Microsoft\Edge\User Data"
    if (Test-Path $edgePath) {
        try {
            $edgeVer = ""
            try { $edgeVer = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Edge\BLBeacon" -ErrorAction Stop).version } catch {}
            $profiles = @(Get-ChildItem -Path $edgePath -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq "Default" -or $_.Name -like "Profile *" })
            $edgeExts = @()
            foreach ($prof in $profiles) { $edgeExts += Get-ChromiumExtensions -ProfileDir $prof.FullName -BrowserName "Edge" }
            $browsers += [ordered]@{
                Name = "Microsoft Edge"
                Version = $edgeVer
                Profiles = @($profiles | ForEach-Object { $_.Name })
                ExtensionCount = $edgeExts.Count
                Extensions = $edgeExts
            }
        } catch { $audit._errors += @{ Section = "BrowserEdge"; Error = $_.Exception.Message } }
    }

    # Chrome
    $chromePath = "$env:LOCALAPPDATA\Google\Chrome\User Data"
    if (Test-Path $chromePath) {
        try {
            $chromeVer = ""
            try { $chromeVer = (Get-ItemProperty "HKLM:\SOFTWARE\Google\Chrome\BLBeacon" -ErrorAction Stop).version } catch {}
            $profiles = @(Get-ChildItem -Path $chromePath -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq "Default" -or $_.Name -like "Profile *" })
            $chromeExts = @()
            foreach ($prof in $profiles) { $chromeExts += Get-ChromiumExtensions -ProfileDir $prof.FullName -BrowserName "Chrome" }
            $browsers += [ordered]@{
                Name = "Google Chrome"
                Version = $chromeVer
                Profiles = @($profiles | ForEach-Object { $_.Name })
                ExtensionCount = $chromeExts.Count
                Extensions = $chromeExts
            }
        } catch { $audit._errors += @{ Section = "BrowserChrome"; Error = $_.Exception.Message } }
    }

    # Brave
    $bravePath = "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data"
    if (Test-Path $bravePath) {
        try {
            $profiles = @(Get-ChildItem -Path $bravePath -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq "Default" -or $_.Name -like "Profile *" })
            $braveExts = @()
            foreach ($prof in $profiles) { $braveExts += Get-ChromiumExtensions -ProfileDir $prof.FullName -BrowserName "Brave" }
            $browsers += [ordered]@{
                Name = "Brave"
                Version = ""
                Profiles = @($profiles | ForEach-Object { $_.Name })
                ExtensionCount = $braveExts.Count
                Extensions = $braveExts
            }
        } catch { $audit._errors += @{ Section = "BrowserBrave"; Error = $_.Exception.Message } }
    }

    # Firefox (different format -- extensions.json)
    $firefoxBase = "$env:APPDATA\Mozilla\Firefox\Profiles"
    if (Test-Path $firefoxBase) {
        try {
            $ffVer = ""
            try { $ffVer = (Get-ItemProperty "HKLM:\SOFTWARE\Mozilla\Mozilla Firefox" -ErrorAction Stop).CurrentVersion } catch {}
            $profiles = @(Get-ChildItem -Path $firefoxBase -Directory -ErrorAction SilentlyContinue)
            $ffExts = @()
            foreach ($prof in $profiles) {
                $extJson = Join-Path $prof.FullName "extensions.json"
                if (Test-Path $extJson) {
                    try {
                        $data = Get-Content $extJson -Raw | ConvertFrom-Json
                        foreach ($a in $data.addons) {
                            if ($a.location -eq "app-system-defaults" -or $a.location -eq "app-builtin") { continue }
                            $ffExts += [ordered]@{
                                Browser = "Firefox"
                                Id = "$($a.id)"
                                Name = "$($a.defaultLocale.name)"
                                Version = "$($a.version)"
                                Active = [bool]$a.active
                                Type = "$($a.type)"
                                SourceURI = "$($a.sourceURI)"
                            }
                        }
                    } catch {}
                }
            }
            $browsers += [ordered]@{
                Name = "Firefox"
                Version = $ffVer
                Profiles = @($profiles | ForEach-Object { $_.Name })
                ExtensionCount = $ffExts.Count
                Extensions = $ffExts
            }
        } catch { $audit._errors += @{ Section = "BrowserFirefox"; Error = $_.Exception.Message } }
    }

    $audit.Browsers = $browsers
    foreach ($b in $browsers) {
        Write-Host "  $($b.Name) $($b.Version): profiles=$($b.Profiles.Count) extensions=$($b.ExtensionCount)"
    }
    # Flag extensions with risky permissions
    $allExts = @()
    foreach ($b in $browsers) { $allExts += $b.Extensions }
    $riskyExts = @($allExts | Where-Object { $_.Permissions -and ($_.Permissions -join ',') -match '(?i)<all_urls>|http\*://|tabs|webRequestBlocking|cookies|history|downloads|management|nativeMessaging|debugger|proxy' })
    if ($riskyExts.Count -gt 0) {
        $audit.SecuritySummary += "$($riskyExts.Count) browser extension(s) with broad/risky permissions"
        Write-Host "  Browser extensions with broad permissions: $($riskyExts.Count)" -ForegroundColor Yellow
    }
} catch {
    Write-Host "  [ERROR] Browser Hygiene section failed: $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "BrowserHygiene"; Error = $_.Exception.Message }
}

# =====================================================
# 44. AUTHENTICATION POSTURE
# =====================================================
Write-Host ""
Write-Host "=== 44. AUTHENTICATION POSTURE ===" -ForegroundColor Cyan
try {
    $authp = [ordered]@{}

    # LSA Protection (RunAsPPL)
    try {
        $runAsPPL = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" -Name RunAsPPL -ErrorAction SilentlyContinue).RunAsPPL
        $runAsPPLBoot = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" -Name RunAsPPLBoot -ErrorAction SilentlyContinue).RunAsPPLBoot
        $authp.LSAProtection = [ordered]@{
            RunAsPPL = $runAsPPL
            RunAsPPLBoot = $runAsPPLBoot
            Enabled = ($runAsPPL -ge 1)
        }
    } catch { $authp.LSAProtection = [ordered]@{ Error = $_.Exception.Message } }

    # Credential Guard / HVCI
    try {
        $dg = Get-CimInstance -ClassName Win32_DeviceGuard -Namespace "root\Microsoft\Windows\DeviceGuard" -ErrorAction SilentlyContinue
        if ($dg) {
            $running = @($dg.SecurityServicesRunning)
            $configured = @($dg.SecurityServicesConfigured)
            $authp.DeviceGuard = [ordered]@{
                CredentialGuardRunning = ($running -contains 1)
                HVCIRunning = ($running -contains 2)
                CredentialGuardConfigured = ($configured -contains 1)
                HVCIConfigured = ($configured -contains 2)
                VirtualizationBasedSecurityStatus = "$($dg.VirtualizationBasedSecurityStatus)"
                CodeIntegrityPolicyEnforcementStatus = "$($dg.CodeIntegrityPolicyEnforcementStatus)"
            }
        } else { $authp.DeviceGuard = [ordered]@{ Available = $false } }
    } catch { $authp.DeviceGuard = [ordered]@{ Error = $_.Exception.Message } }

    # WDigest
    try {
        $wdigest = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest" -Name UseLogonCredential -ErrorAction SilentlyContinue).UseLogonCredential
        $authp.WDigest = [ordered]@{
            UseLogonCredential = $wdigest
            CleartextDisabled = ($wdigest -eq $null -or $wdigest -eq 0)
        }
    } catch { $authp.WDigest = [ordered]@{ Error = $_.Exception.Message } }

    # NTLM / LM
    try {
        $lmCompat = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" -Name LmCompatibilityLevel -ErrorAction SilentlyContinue).LmCompatibilityLevel
        $noLMHash = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" -Name NoLMHash -ErrorAction SilentlyContinue).NoLMHash
        $authp.NTLM = [ordered]@{
            LmCompatibilityLevel = $lmCompat
            NoLMHash = $noLMHash
            LMHashStored = ($noLMHash -ne 1)
        }
    } catch { $authp.NTLM = [ordered]@{ Error = $_.Exception.Message } }

    # Cached creds
    try {
        $cached = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" -Name CachedLogonsCount -ErrorAction SilentlyContinue).CachedLogonsCount
        $authp.CachedLogonsCount = $cached
    } catch { $authp.CachedLogonsCount = $null }

    # klist (only meaningful for current session, may be empty when run as SYSTEM)
    try {
        $klistOut = (& klist 2>&1) -join "`n"
        $tickets = @([regex]::Matches($klistOut, '#\d+>'))
        $authp.KerberosTicketCount = $tickets.Count
    } catch { $authp.KerberosTicketCount = $null }

    # dsregcmd /status -- AzureAD/Hybrid join state + WHfB
    try {
        $dsreg = (& dsregcmd /status 2>&1) -join "`n"
        $authp.DSRegStatus = [ordered]@{
            AzureAdJoined = ($dsreg -match 'AzureAdJoined\s*:\s*YES')
            DomainJoined = ($dsreg -match 'DomainJoined\s*:\s*YES')
            EnterpriseJoined = ($dsreg -match 'EnterpriseJoined\s*:\s*YES')
            DeviceId = if ($dsreg -match 'DeviceId\s*:\s*([\w-]+)') { $matches[1] } else { "" }
            TenantName = if ($dsreg -match 'TenantName\s*:\s*(.+)') { $matches[1].Trim() } else { "" }
            TenantId = if ($dsreg -match 'TenantId\s*:\s*([\w-]+)') { $matches[1] } else { "" }
            WHfBEnabled = ($dsreg -match 'WamDefaultGUID.+microsoft' -or $dsreg -match 'NgcSet\s*:\s*YES')
        }
    } catch { $authp.DSRegStatus = [ordered]@{ Error = $_.Exception.Message } }

    $audit.AuthenticationPosture = $authp

    Write-Host "  LSA Protection (RunAsPPL): $($authp.LSAProtection.Enabled)"
    Write-Host "  Credential Guard running: $($authp.DeviceGuard.CredentialGuardRunning)"
    Write-Host "  WDigest cleartext disabled: $($authp.WDigest.CleartextDisabled)"
    Write-Host "  NTLM LmCompatibilityLevel: $($authp.NTLM.LmCompatibilityLevel) (5 = recommended)"
    Write-Host "  Cached logons count: $($authp.CachedLogonsCount)"
    Write-Host "  AzureAdJoined: $($authp.DSRegStatus.AzureAdJoined), DomainJoined: $($authp.DSRegStatus.DomainJoined)"

    # Findings
    if ($authp.LSAProtection.Enabled -eq $false) {
        $audit.SecuritySummary += "LSA Protection (RunAsPPL) not enabled"
    }
    if ($authp.WDigest.CleartextDisabled -eq $false) {
        $audit.SecuritySummary += "WDigest cleartext credentials not disabled"
    }
    if ($authp.NTLM.LmCompatibilityLevel -ne $null -and $authp.NTLM.LmCompatibilityLevel -lt 5) {
        $audit.SecuritySummary += "NTLM LmCompatibilityLevel = $($authp.NTLM.LmCompatibilityLevel) (recommend 5)"
    }
    if ($authp.NTLM.LMHashStored) {
        $audit.SecuritySummary += "LM hashes still stored (NoLMHash != 1)"
    }
    if ($authp.CachedLogonsCount -ne $null -and $authp.CachedLogonsCount -gt 4) {
        $audit.SecuritySummary += "CachedLogonsCount = $($authp.CachedLogonsCount) (recommend <= 4 for security)"
    }
} catch {
    Write-Host "  [ERROR] Authentication Posture section failed: $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "AuthenticationPosture"; Error = $_.Exception.Message }
}

# =====================================================
# 45. HARDWARE DEEPER (battery, SMART, driver problems)
# =====================================================
Write-Host ""
Write-Host "=== 45. HARDWARE DEEPER ===" -ForegroundColor Cyan
try {
    $hw = [ordered]@{}

    # Battery (laptop)
    try {
        $bat = @(Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue)
        if ($bat.Count -gt 0) {
            $batStatic = @(Get-CimInstance -Namespace "root\wmi" -ClassName BatteryStaticData -ErrorAction SilentlyContinue)
            $batFull = @(Get-CimInstance -Namespace "root\wmi" -ClassName BatteryFullChargedCapacity -ErrorAction SilentlyContinue)
            $batCycle = @(Get-CimInstance -Namespace "root\wmi" -ClassName BatteryCycleCount -ErrorAction SilentlyContinue)
            $hw.Batteries = @()
            for ($i=0; $i -lt $bat.Count; $i++) {
                $design = if ($i -lt $batStatic.Count) { $batStatic[$i].DesignedCapacity } else { $null }
                $full = if ($i -lt $batFull.Count) { $batFull[$i].FullChargedCapacity } else { $null }
                $cycles = if ($i -lt $batCycle.Count) { $batCycle[$i].CycleCount } else { $null }
                $healthPct = if ($design -and $full -and $design -gt 0) { [math]::Round(($full / $design) * 100, 1) } else { $null }
                $hw.Batteries += [ordered]@{
                    Name = "$($bat[$i].Name)"
                    Status = "$($bat[$i].Status)"
                    BatteryStatusCode = $bat[$i].BatteryStatus
                    EstimatedChargeRemainingPct = $bat[$i].EstimatedChargeRemaining
                    DesignCapacity_mWh = $design
                    FullChargeCapacity_mWh = $full
                    CycleCount = $cycles
                    HealthPercent = $healthPct
                }
                if ($healthPct -ne $null -and $healthPct -lt 60) {
                    $audit.SecuritySummary += "Battery health $healthPct% (consider replacement)"
                }
            }
        } else {
            $hw.Batteries = @()
        }
    } catch { $hw.BatteriesError = $_.Exception.Message; $hw.Batteries = @() }

    # SMART per physical disk
    try {
        $physDisks = @(Get-PhysicalDisk -ErrorAction SilentlyContinue)
        $hw.SMART = @($physDisks | ForEach-Object {
            $d = $_
            $rel = $null
            try { $rel = $d | Get-StorageReliabilityCounter -ErrorAction SilentlyContinue } catch {}
            [ordered]@{
                FriendlyName = "$($d.FriendlyName)"
                MediaType = "$($d.MediaType)"
                BusType = "$($d.BusType)"
                HealthStatus = "$($d.HealthStatus)"
                OperationalStatus = "$($d.OperationalStatus)"
                SizeGB = if ($d.Size) { [math]::Round($d.Size/1GB, 1) } else { $null }
                Wear = if ($rel) { $rel.Wear } else { $null }
                Temperature = if ($rel) { $rel.Temperature } else { $null }
                ReadErrorsTotal = if ($rel) { $rel.ReadErrorsTotal } else { $null }
                WriteErrorsTotal = if ($rel) { $rel.WriteErrorsTotal } else { $null }
                PowerOnHours = if ($rel) { $rel.PowerOnHours } else { $null }
                StartStopCycleCount = if ($rel) { $rel.StartStopCycleCount } else { $null }
            }
        })
        $unhealthy = @($hw.SMART | Where-Object { $_.HealthStatus -ne "Healthy" })
        if ($unhealthy.Count -gt 0) {
            foreach ($d in $unhealthy) {
                $audit.SecuritySummary += "Disk SMART unhealthy: $($d.FriendlyName) ($($d.HealthStatus))"
            }
        }
    } catch { $hw.SMARTError = $_.Exception.Message; $hw.SMART = @() }

    # Driver problems (yellow-bang devices)
    try {
        if (Get-Command Get-PnpDevice -ErrorAction SilentlyContinue) {
            $problemDevs = @(Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue | Where-Object { $_.Status -in 'Error','Degraded','Unknown' })
            $hw.ProblemDevices = @($problemDevs | ForEach-Object {
                [ordered]@{
                    FriendlyName = $_.FriendlyName
                    Class = "$($_.Class)"
                    Status = "$($_.Status)"
                    Manufacturer = "$($_.Manufacturer)"
                    InstanceId = $_.InstanceId
                    ProblemCode = "$($_.Problem)"
                }
            })
            if ($problemDevs.Count -gt 0) {
                $audit.SecuritySummary += "$($problemDevs.Count) device(s) with driver/PNP errors"
            }
        }
    } catch { $hw.ProblemDevicesError = $_.Exception.Message; $hw.ProblemDevices = @() }

    $audit.HardwareDeeper = $hw
    Write-Host "  Batteries: $(@($hw.Batteries).Count) -- worst health: $((@($hw.Batteries | ForEach-Object HealthPercent | Where-Object { $_ -ne $null } | Sort-Object | Select-Object -First 1)))%"
    Write-Host "  Physical disks: $(@($hw.SMART).Count) -- unhealthy: $(@($hw.SMART | Where-Object { $_.HealthStatus -ne 'Healthy' }).Count)"
    Write-Host "  Devices with driver errors: $(@($hw.ProblemDevices).Count)"
} catch {
    Write-Host "  [ERROR] Hardware Deeper section failed: $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "HardwareDeeper"; Error = $_.Exception.Message }
}

# =====================================================
# 46. PERFORMANCE SNAPSHOT (top procs, memory, page file, uptime)
# =====================================================
Write-Host ""
Write-Host "=== 46. PERFORMANCE SNAPSHOT ===" -ForegroundColor Cyan
try {
    $perf = [ordered]@{}

    # Top processes
    try {
        $procs = @(Get-Process -ErrorAction SilentlyContinue)
        $perf.TopByCPU = @($procs | Where-Object { $_.CPU -ne $null } | Sort-Object CPU -Descending | Select-Object -First 10 | ForEach-Object {
            [ordered]@{
                Name = $_.ProcessName
                Id = $_.Id
                CPUSeconds = [math]::Round($_.CPU, 1)
                WorkingSetMB = [math]::Round($_.WS / 1MB, 1)
                Path = if ($_.Path) { $_.Path } else { "" }
            }
        })
        $perf.TopByMemory = @($procs | Sort-Object WS -Descending | Select-Object -First 10 | ForEach-Object {
            [ordered]@{
                Name = $_.ProcessName
                Id = $_.Id
                WorkingSetMB = [math]::Round($_.WS / 1MB, 1)
                CPUSeconds = if ($_.CPU) { [math]::Round($_.CPU, 1) } else { 0 }
                Path = if ($_.Path) { $_.Path } else { "" }
            }
        })
        $perf.ProcessTotalCount = $procs.Count
    } catch { $perf.TopProcessError = $_.Exception.Message }

    # Memory
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $perf.Memory = [ordered]@{
            TotalGB = [math]::Round($os.TotalVisibleMemorySize / 1MB, 1)
            FreeGB = [math]::Round($os.FreePhysicalMemory / 1MB, 1)
            UsedPct = [math]::Round((($os.TotalVisibleMemorySize - $os.FreePhysicalMemory) / $os.TotalVisibleMemorySize) * 100, 1)
        }
        $perf.Uptime = [ordered]@{
            LastBootTime = $os.LastBootUpTime.ToString("yyyy-MM-dd HH:mm:ss")
            UptimeDays = [math]::Round(((Get-Date) - $os.LastBootUpTime).TotalDays, 2)
        }
        if ($perf.Uptime.UptimeDays -gt 30) {
            $audit.SecuritySummary += "System uptime $($perf.Uptime.UptimeDays) days (recommend reboot for patches)"
        }
        if ($perf.Memory.UsedPct -gt 90) {
            $audit.SecuritySummary += "Memory used $($perf.Memory.UsedPct)% (high pressure)"
        }
    } catch { $perf.MemoryError = $_.Exception.Message }

    # Page file
    try {
        $pf = @(Get-CimInstance -ClassName Win32_PageFileUsage -ErrorAction SilentlyContinue)
        $perf.PageFiles = @($pf | ForEach-Object {
            [ordered]@{
                Name = $_.Name
                AllocatedBaseSizeMB = $_.AllocatedBaseSize
                CurrentUsageMB = $_.CurrentUsage
                PeakUsageMB = $_.PeakUsage
            }
        })
    } catch { $perf.PageFiles = @() }

    # Recent reboots (last 10)
    try {
        $rebootEvents = @(Get-WinEvent -FilterHashtable @{LogName='System'; Id=1074,6005,6006,6008,41} -MaxEvents 50 -ErrorAction SilentlyContinue)
        $perf.RecentReboots = @($rebootEvents | Sort-Object TimeCreated -Descending | Select-Object -First 10 | ForEach-Object {
            [ordered]@{
                Time = $_.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
                Id = $_.Id
                ProviderName = $_.ProviderName
                Message = (($_.Message -split "`r?`n")[0]).Trim()
            }
        })
    } catch { $perf.RecentReboots = @() }

    $audit.Performance = $perf
    Write-Host "  Processes: $($perf.ProcessTotalCount), Memory: $($perf.Memory.UsedPct)% used ($($perf.Memory.FreeGB)/$($perf.Memory.TotalGB) GB free)"
    Write-Host "  Uptime: $($perf.Uptime.UptimeDays) days (last boot: $($perf.Uptime.LastBootTime))"
    Write-Host "  Page files: $($perf.PageFiles.Count), Recent reboot events: $($perf.RecentReboots.Count)"
} catch {
    Write-Host "  [ERROR] Performance section failed: $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "Performance"; Error = $_.Exception.Message }
}

# =====================================================
# 47. OFFICE / OUTLOOK
# =====================================================
Write-Host ""
Write-Host "=== 47. OFFICE / OUTLOOK ===" -ForegroundColor Cyan
try {
    $office = [ordered]@{}

    # Office C2R config
    try {
        $c2r = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue
        if ($c2r) {
            $office.ClickToRun = [ordered]@{
                VersionToReport = "$($c2r.VersionToReport)"
                ProductReleaseIds = "$($c2r.ProductReleaseIds)"
                CDNBaseUrl = "$($c2r.CDNBaseUrl)"
                UpdateChannel = "$($c2r.UpdateChannel)"
                ClientCulture = "$($c2r.ClientCulture)"
                Platform = "$($c2r.Platform)"
                SharedComputerLicensing = $c2r.SharedComputerLicensing
            }
        }
    } catch { $office.ClickToRunError = $_.Exception.Message }

    # Outlook profiles + accounts
    try {
        $outlookProfileBases = @(
            "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles",
            "HKCU:\Software\Microsoft\Office\15.0\Outlook\Profiles"
        )
        $office.OutlookProfiles = @()
        foreach ($base in $outlookProfileBases) {
            if (Test-Path $base) {
                Get-ChildItem -Path $base -ErrorAction SilentlyContinue | ForEach-Object {
                    $office.OutlookProfiles += [ordered]@{
                        Hive = $base
                        ProfileName = $_.PSChildName
                    }
                }
            }
        }
    } catch { $office.OutlookProfilesError = $_.Exception.Message }

    # OST/PST sizes
    try {
        $pstLocations = @(
            "$env:LOCALAPPDATA\Microsoft\Outlook",
            "$env:USERPROFILE\Documents\Outlook Files"
        ) | Where-Object { Test-Path $_ }
        $office.MailFiles = @()
        foreach ($loc in $pstLocations) {
            Get-ChildItem -Path $loc -Filter "*.ost" -File -ErrorAction SilentlyContinue | ForEach-Object {
                $office.MailFiles += [ordered]@{ Type = "OST"; Path = $_.FullName; SizeGB = [math]::Round($_.Length / 1GB, 2); LastWrite = $_.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss") }
            }
            Get-ChildItem -Path $loc -Filter "*.pst" -File -ErrorAction SilentlyContinue | ForEach-Object {
                $office.MailFiles += [ordered]@{ Type = "PST"; Path = $_.FullName; SizeGB = [math]::Round($_.Length / 1GB, 2); LastWrite = $_.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss") }
            }
        }
        $largeOST = @($office.MailFiles | Where-Object { $_.SizeGB -gt 50 })
        if ($largeOST.Count -gt 0) {
            $audit.SecuritySummary += "$($largeOST.Count) large Outlook data file(s) >50GB (sync/perf risk)"
        }
    } catch { $office.MailFilesError = $_.Exception.Message }

    # Outlook add-ins
    try {
        $office.Addins = @()
        foreach ($base in @("HKCU:\Software\Microsoft\Office\Outlook\Addins","HKLM:\Software\Microsoft\Office\Outlook\Addins","HKLM:\Software\Wow6432Node\Microsoft\Office\Outlook\Addins")) {
            if (Test-Path $base) {
                Get-ChildItem -Path $base -ErrorAction SilentlyContinue | ForEach-Object {
                    $sub = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
                    $office.Addins += [ordered]@{
                        Hive = $base
                        ProgId = $_.PSChildName
                        FriendlyName = "$($sub.FriendlyName)"
                        Description = "$($sub.Description)"
                        LoadBehavior = $sub.LoadBehavior
                    }
                }
            }
        }
    } catch { $office.AddinsError = $_.Exception.Message }

    $audit.OfficeOutlook = $office
    Write-Host "  Office C2R version: $($office.ClickToRun.VersionToReport), channel: $($office.ClickToRun.UpdateChannel)"
    Write-Host "  Outlook profiles: $($office.OutlookProfiles.Count), Mail files: $($office.MailFiles.Count), Add-ins: $($office.Addins.Count)"
} catch {
    Write-Host "  [ERROR] Office/Outlook section failed: $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "OfficeOutlook"; Error = $_.Exception.Message }
}

# =====================================================
# 48. TIME + UPDATE EXTENDED
# =====================================================
Write-Host ""
Write-Host "=== 48. TIME / UPDATE EXTENDED ===" -ForegroundColor Cyan
try {
    $tu = [ordered]@{}

    # Time service
    try {
        $w32tm = (& w32tm /query /status 2>&1) -join "`n"
        $tu.W32time = [ordered]@{
            Source = if ($w32tm -match 'Source:\s+(.+)') { $matches[1].Trim() } else { "" }
            LastSync = if ($w32tm -match 'Last Successful Sync Time:\s+(.+)') { $matches[1].Trim() } else { "" }
            Stratum = if ($w32tm -match 'Stratum:\s+(\d+)') { $matches[1] } else { "" }
            LeapIndicator = if ($w32tm -match 'Leap Indicator:\s+(.+)') { $matches[1].Trim() } else { "" }
        }
    } catch { $tu.W32time = [ordered]@{ Error = $_.Exception.Message } }

    # Time zone
    try {
        $tz = Get-TimeZone -ErrorAction SilentlyContinue
        $tu.TimeZone = if ($tz) { [ordered]@{ Id = $tz.Id; DisplayName = $tz.DisplayName; BaseUtcOffset = "$($tz.BaseUtcOffset)" } } else { @{} }
    } catch { $tu.TimeZone = @{} }

    # Pending Windows Updates (COM)
    try {
        $session = New-Object -ComObject Microsoft.Update.Session -ErrorAction Stop
        $searcher = $session.CreateUpdateSearcher()
        $searchResult = $searcher.Search("IsInstalled=0 and IsHidden=0 and Type='Software'")
        $tu.PendingUpdates = [ordered]@{
            Count = $searchResult.Updates.Count
            Updates = @()
        }
        for ($i=0; $i -lt [math]::Min($searchResult.Updates.Count, 25); $i++) {
            $up = $searchResult.Updates.Item($i)
            $tu.PendingUpdates.Updates += [ordered]@{
                Title = "$($up.Title)"
                IsCritical = ($up.MsrcSeverity -eq "Critical")
                Severity = "$($up.MsrcSeverity)"
                KBs = @($up.KBArticleIDs)
                SizeMB = if ($up.MaxDownloadSize) { [math]::Round($up.MaxDownloadSize / 1MB, 1) } else { $null }
            }
        }
        if ($searchResult.Updates.Count -gt 0) {
            $audit.SecuritySummary += "$($searchResult.Updates.Count) pending Windows Update(s)"
        }
    } catch { $tu.PendingUpdates = [ordered]@{ Error = $_.Exception.Message } }

    # WU history failures (last 30d)
    try {
        $wu30 = (Get-Date).AddDays(-30)
        $failures = @(Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-WindowsUpdateClient/Operational'; Id=20,25,31; StartTime=$wu30} -ErrorAction SilentlyContinue)
        $tu.WUHistoryFailures30d = [ordered]@{
            Count = $failures.Count
            Last10 = @($failures | Sort-Object TimeCreated -Descending | Select-Object -First 10 | ForEach-Object {
                [ordered]@{
                    Time = $_.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
                    Id = $_.Id
                    Message = (($_.Message -split "`r?`n")[0]).Trim()
                }
            })
        }
    } catch { $tu.WUHistoryFailures30d = [ordered]@{ Error = $_.Exception.Message } }

    $audit.TimeUpdateExtended = $tu
    Write-Host "  Time source: $($tu.W32time.Source) | Last sync: $($tu.W32time.LastSync)"
    Write-Host "  Time zone: $($tu.TimeZone.Id)"
    Write-Host "  Pending updates: $($tu.PendingUpdates.Count) | WU failures (30d): $($tu.WUHistoryFailures30d.Count)"
} catch {
    Write-Host "  [ERROR] Time/Update extended section failed: $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "TimeUpdateExtended"; Error = $_.Exception.Message }
}

# =====================================================
# 49. EVENT LOG ADVANCED (crashes, lockouts, encoded PS, defender)
# =====================================================
Write-Host ""
Write-Host "=== 49. EVENT LOG ADVANCED ===" -ForegroundColor Cyan
try {
    $ev = [ordered]@{}
    $thirtyDays = (Get-Date).AddDays(-30)

    # Top crashing apps (Application 1000)
    try {
        $crashes = @(Get-WinEvent -FilterHashtable @{LogName='Application'; ProviderName='Application Error'; Id=1000; StartTime=$thirtyDays} -ErrorAction SilentlyContinue)
        $crashGrouped = $crashes | ForEach-Object {
            # Message format: "Faulting application name: <name>, version..."
            if ($_.Message -match 'Faulting application name:\s+(\S+)') { $matches[1] } else { "unknown" }
        } | Group-Object | Sort-Object Count -Descending | Select-Object -First 10
        $ev.AppCrashesTop10 = @($crashGrouped | ForEach-Object {
            [ordered]@{ Application = $_.Name; CrashCount = $_.Count }
        })
        $ev.AppCrashesTotal = $crashes.Count
    } catch { $ev.AppCrashesError = $_.Exception.Message }

    # Account lockouts (Security 4740)
    try {
        $lockouts = @(Get-WinEvent -FilterHashtable @{LogName='Security'; Id=4740; StartTime=$thirtyDays} -ErrorAction SilentlyContinue)
        $ev.AccountLockouts30d = @($lockouts | Sort-Object TimeCreated -Descending | Select-Object -First 25 | ForEach-Object {
            $msg = $_.Message
            $accountName = if ($msg -match 'Account That Was Locked Out:\s*\r?\n\s*Security ID:\s+\S+\s*\r?\n\s*Account Name:\s+(.+)') { $matches[1].Trim() } else { "" }
            $callerComputer = if ($msg -match 'Caller Computer Name:\s+(.+)') { $matches[1].Trim() } else { "" }
            [ordered]@{
                Time = $_.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
                AccountName = $accountName
                CallerComputer = $callerComputer
            }
        })
        if ($lockouts.Count -gt 0) {
            $audit.SecuritySummary += "$($lockouts.Count) account lockout event(s) in last 30d"
        }
    } catch { $ev.AccountLockoutsError = $_.Exception.Message }

    # Audit policy changes (Security 4719)
    try {
        $polChanges = @(Get-WinEvent -FilterHashtable @{LogName='Security'; Id=4719; StartTime=$thirtyDays} -ErrorAction SilentlyContinue)
        $ev.AuditPolicyChanges30d = $polChanges.Count
        if ($polChanges.Count -gt 0) {
            $audit.SecuritySummary += "$($polChanges.Count) audit policy change event(s) in last 30d"
        }
    } catch { $ev.AuditPolicyChangesError = $_.Exception.Message }

    # PowerShell encoded commands (Microsoft-Windows-PowerShell/Operational, EID 4104)
    try {
        $encScripts = @(Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-PowerShell/Operational'; Id=4104; StartTime=$thirtyDays} -MaxEvents 500 -ErrorAction SilentlyContinue |
            Where-Object {
                $m = $_.Message
                ($m -match 'FromBase64String' -or $m -match '-EncodedCommand' -or $m -match '-enc\s+[A-Za-z0-9+/=]{50,}' -or $m -match 'IEX\s*\(' -or $m -match 'Invoke-Expression')
            })
        $ev.SuspiciousPowerShell30d = [ordered]@{
            Count = $encScripts.Count
            Last10 = @($encScripts | Sort-Object TimeCreated -Descending | Select-Object -First 10 | ForEach-Object {
                $snippet = ($_.Message -split "`r?`n" | Select-Object -First 3) -join " "
                if ($snippet.Length -gt 300) { $snippet = $snippet.Substring(0, 300) + "..." }
                [ordered]@{
                    Time = $_.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
                    Snippet = $snippet
                }
            })
        }
        if ($encScripts.Count -gt 5) {
            $audit.SecuritySummary += "$($encScripts.Count) suspicious PowerShell scripts (base64/IEX/encoded) in last 30d"
        }
    } catch { $ev.SuspiciousPowerShellError = $_.Exception.Message }

    # Defender detections (Microsoft-Windows-Windows Defender/Operational, EID 1116, 1117)
    try {
        $defDet = @(Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Windows Defender/Operational'; Id=1116,1117,1118,1119; StartTime=$thirtyDays} -ErrorAction SilentlyContinue)
        $ev.DefenderDetections30d = [ordered]@{
            Count = $defDet.Count
            Last10 = @($defDet | Sort-Object TimeCreated -Descending | Select-Object -First 10 | ForEach-Object {
                [ordered]@{
                    Time = $_.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
                    Id = $_.Id
                    Message = (($_.Message -split "`r?`n")[0]).Trim()
                }
            })
        }
        if ($defDet.Count -gt 0) {
            $audit.SecuritySummary += "$($defDet.Count) Defender detection event(s) in last 30d"
        }
    } catch { $ev.DefenderDetectionsError = $_.Exception.Message }

    # Sysmon (if installed, EID 1 = process create) - flag unsigned in user dirs
    try {
        $sysmon = @(Get-WinEvent -ListLog "Microsoft-Windows-Sysmon/Operational" -ErrorAction SilentlyContinue)
        if ($sysmon) {
            $ev.SysmonInstalled = $true
            $sysmonProc = @(Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Sysmon/Operational'; Id=1; StartTime=$thirtyDays} -MaxEvents 200 -ErrorAction SilentlyContinue |
                Where-Object {
                    $m = $_.Message
                    ($m -match 'Image:\s+(C:\\Users\\[^\r\n]+\.exe)' -or $m -match 'Image:\s+(C:\\ProgramData\\[^\r\n]+\.exe)' -or $m -match 'Image:\s+(C:\\Windows\\Temp\\[^\r\n]+\.exe)') -and
                    ($m -notmatch 'Signed:\s+true')
                })
            $ev.SysmonUnsignedExecsInUserDirs = $sysmonProc.Count
            if ($sysmonProc.Count -gt 0) {
                $audit.SecuritySummary += "Sysmon: $($sysmonProc.Count) unsigned exec(s) from user-writable dirs in 30d"
            }
        } else { $ev.SysmonInstalled = $false }
    } catch { $ev.SysmonError = $_.Exception.Message }

    $audit.EventLogAdvanced = $ev
    Write-Host "  App crashes (30d): $($ev.AppCrashesTotal) total, top 10 apps captured"
    Write-Host "  Account lockouts (30d): $(@($ev.AccountLockouts30d).Count)"
    Write-Host "  Suspicious PowerShell (30d): $($ev.SuspiciousPowerShell30d.Count)"
    Write-Host "  Defender detections (30d): $($ev.DefenderDetections30d.Count)"
    Write-Host "  Sysmon installed: $($ev.SysmonInstalled)"
} catch {
    Write-Host "  [ERROR] Event Log Advanced section failed: $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "EventLogAdvanced"; Error = $_.Exception.Message }
}

# =====================================================
# DONE - SAVE JSON
# =====================================================
Write-Host ""
Write-Host "======================================="
Write-Host "    WORKSTATION AUDIT COMPLETE"
Write-Host "======================================="

$errorCount = $audit._errors.Count
if ($errorCount -gt 0) {
    Write-Host "Completed with $errorCount section errors (see _errors in JSON)" -ForegroundColor Yellow
} else {
    Write-Host "All sections completed successfully" -ForegroundColor Green
}

Write-Host "JSON data: $JsonFile"
Write-Host "======================================="

# Save JSON
$audit | ConvertTo-Json -Depth 10 | Out-File $JsonFile -Encoding UTF8
