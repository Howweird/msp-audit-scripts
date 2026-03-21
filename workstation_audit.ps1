# ==========================================
# UNIVERSAL WORKSTATION AUDIT SCRIPT
# Works on Windows 10 / 11 (domain or workgroup)
# Outputs: HOSTNAME_workstation_audit_DATE.json to C:\Temp
# ==========================================

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

# Structured data collector
$audit = [ordered]@{
    _metadata = [ordered]@{
        ScriptVersion = "1.0"
        ScriptType    = "Workstation"
        RunDate       = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        RunBy         = "$env:USERDOMAIN\$env:USERNAME"
        Hostname      = $env:COMPUTERNAME
    }
    _errors = @()
}

Write-Host "======================================="
Write-Host "    UNIVERSAL WORKSTATION AUDIT v1.0"
Write-Host "    $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "======================================="
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
