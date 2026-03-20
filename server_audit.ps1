# ==========================================
# UNIVERSAL WINDOWS SERVER AUDIT SCRIPT
# Works on Windows Server 2012 -> 2025
# Outputs: .txt transcript + .json for AI processing
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
$TxtFile = "$OutputDir\server_audit_$Date.txt"
$JsonFile = "$OutputDir\server_audit_$Date.json"

if (!(Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }

# Prevent transcript truncation
$FormatEnumerationLimit = -1
try { $Host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size(500, 50000) } catch {}

# Structured data collector
$audit = [ordered]@{
    _metadata = [ordered]@{
        ScriptVersion = "2.1"
        ScriptType    = "Server"
        RunDate       = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        RunBy         = "$env:USERDOMAIN\$env:USERNAME"
        Hostname      = $env:COMPUTERNAME
    }
    _errors = @()
}

Start-Transcript -Path $TxtFile -Force

Write-Host "======================================="
Write-Host "      UNIVERSAL SERVER AUDIT v2.0"
Write-Host "      $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"
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

    $sysInfo = [ordered]@{
        Hostname      = $env:COMPUTERNAME
        OS            = $os.Caption
        OSVersion     = $os.Version
        BuildNumber   = $os.BuildNumber
        Architecture  = $os.OSArchitecture
        InstallDate   = $os.InstallDate.ToString("yyyy-MM-dd")
        LastBoot      = $os.LastBootUpTime.ToString("yyyy-MM-dd HH:mm:ss")
        UptimeDays    = [math]::Round($uptime.TotalDays, 1)
        Domain        = $cs.Domain
        DomainJoined  = $cs.PartOfDomain
        DomainRole    = switch ($cs.DomainRole) {
            0 {"Standalone Workstation"} 1 {"Member Workstation"} 2 {"Standalone Server"}
            3 {"Member Server"} 4 {"Backup DC"} 5 {"Primary DC"} default {"Unknown ($($cs.DomainRole))"}
        }
        TotalRAM_GB   = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
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

    $hardware = [ordered]@{
        Manufacturer  = $hw.Manufacturer
        Model         = $hw.Model
        SerialNumber  = if ($bios.SerialNumber) { $bios.SerialNumber } else { $encl.SerialNumber }
        CPU           = $cpu.Name
        CPUCores      = $cpu.NumberOfCores
        CPULogical    = $cpu.NumberOfLogicalProcessors
        RAM_GB        = [math]::Round($hw.TotalPhysicalMemory / 1GB, 1)
        BIOSVersion   = $bios.SMBIOSBIOSVersion
        BIOSDate      = if ($bios.ReleaseDate) { $bios.ReleaseDate.ToString("yyyy-MM-dd") } else { "N/A" }
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
# 3. STORAGE
# =====================================================
Write-Host ""
Write-Host "=== 3. STORAGE ===" -ForegroundColor Cyan

Write-Host "  Running: Get-Volume"
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
            Drive = "$($_.DriveLetter):"
            Label = $_.FileSystemLabel
            FileSystem = $_.FileSystem
            SizeGB = $_.SizeGB
            FreeGB = $_.FreeGB
            UsedPct = $_.UsedPct
            Health = "$($_.HealthStatus)"
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
            DeviceId = "$($_.DeviceId)"
            Name = $_.FriendlyName
            MediaType = "$($_.MediaType)"
            SizeGB = $_.SizeGB
            Health = "$($_.HealthStatus)"
            Status = "$($_.OperationalStatus)"
        }
    })
    Write-Host "  [OK] Physical disks collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "PhysicalDisks"; Error = $_.Exception.Message }
}

# =====================================================
# 4. NETWORK CONFIG
# =====================================================
Write-Host ""
Write-Host "=== 4. NETWORK CONFIG ===" -ForegroundColor Cyan

Write-Host "  Running: Get-NetAdapter"
try {
    $adapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } |
        Select-Object Name, InterfaceDescription, MacAddress, Status, LinkSpeed
    $adapters | Format-Table -AutoSize
    $audit.NetworkAdapters = @($adapters | ForEach-Object {
        [ordered]@{
            Name = $_.Name
            Description = $_.InterfaceDescription
            MAC = $_.MacAddress
            Status = "$($_.Status)"
            LinkSpeed = $_.LinkSpeed
        }
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
            Gateway = "$($_.IPv4DefaultGateway.NextHop)"
            DNS = @($_.DNSServer | Where-Object { $_.AddressFamily -eq 2 } | ForEach-Object { $_.ServerAddresses }) | Select-Object -Unique
        }
    })
    $ipConfigs | Format-List InterfaceAlias, IPv4Address, IPv4DefaultGateway, DnsServer
    Write-Host "  [OK] IP config collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "IPConfig"; Error = $_.Exception.Message }
}

Write-Host "  Running: Get-DnsClientServerAddress"
try {
    Get-DnsClientServerAddress -AddressFamily IPv4 | Where-Object { $_.ServerAddresses.Count -gt 0 } |
        Format-Table InterfaceAlias, ServerAddresses -AutoSize
    Write-Host "  [OK] DNS client config collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "DnsClient"; Error = $_.Exception.Message }
}

# =====================================================
# 5. ROUTING TABLE
# =====================================================
Write-Host ""
Write-Host "=== 5. ROUTING TABLE ===" -ForegroundColor Cyan
Write-Host "  Running: Get-NetRoute"
try {
    Get-NetRoute -AddressFamily IPv4 | Where-Object { $_.DestinationPrefix -ne '255.255.255.255/32' } |
        Sort-Object DestinationPrefix |
        Format-Table DestinationPrefix, NextHop, InterfaceAlias, RouteMetric -AutoSize
    Write-Host "  [OK] Routes collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "Routes"; Error = $_.Exception.Message }
}

# =====================================================
# 6. ARP TABLE
# =====================================================
Write-Host ""
Write-Host "=== 6. ARP TABLE ===" -ForegroundColor Cyan
Write-Host "  Running: Get-NetNeighbor"
try {
    $arp = Get-NetNeighbor -AddressFamily IPv4 | Where-Object { $_.State -ne 'Unreachable' -and $_.State -ne 'Permanent' } |
        Select-Object IPAddress, LinkLayerAddress, State, InterfaceAlias
    $arp | Format-Table -AutoSize
    $audit.ARPTable = @($arp | ForEach-Object {
        [ordered]@{ IP = $_.IPAddress; MAC = $_.LinkLayerAddress; State = "$($_.State)"; Interface = $_.InterfaceAlias }
    })
    Write-Host "  [OK] ARP table collected - $($arp.Count) entries" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "ARP"; Error = $_.Exception.Message }
}

# =====================================================
# 7. LISTENING PORTS
# =====================================================
Write-Host ""
Write-Host "=== 7. LISTENING PORTS ===" -ForegroundColor Cyan
Write-Host "  Running: Get-NetTCPConnection -State Listen"
try {
    $listeners = Get-NetTCPConnection -State Listen | ForEach-Object {
        $proc = Get-Process -Id $_.OwningProcess -ErrorAction SilentlyContinue
        [PSCustomObject]@{
            LocalAddress = $_.LocalAddress
            LocalPort    = $_.LocalPort
            PID          = $_.OwningProcess
            ProcessName  = $proc.ProcessName
        }
    } | Sort-Object LocalPort
    $listeners | Format-Table -AutoSize
    $audit.ListeningPorts = @($listeners | ForEach-Object {
        [ordered]@{ Address = $_.LocalAddress; Port = $_.LocalPort; PID = $_.PID; Process = $_.ProcessName }
    })
    Write-Host "  [OK] Listening ports collected - $($listeners.Count) ports" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "ListeningPorts"; Error = $_.Exception.Message }
}

# =====================================================
# 8. WINDOWS FIREWALL RULES
# =====================================================
Write-Host ""
Write-Host "=== 8. WINDOWS FIREWALL RULES (Enabled) ===" -ForegroundColor Cyan
Write-Host "  Running: Get-NetFirewallRule | Get-NetFirewallPortFilter"
try {
    $fwRules = Get-NetFirewallRule | Where-Object { $_.Enabled -eq 'True' } | ForEach-Object {
        $portFilter = $_ | Get-NetFirewallPortFilter -ErrorAction SilentlyContinue
        [PSCustomObject]@{
            DisplayName = $_.DisplayName
            Direction   = "$($_.Direction)"
            Action      = "$($_.Action)"
            Protocol    = $portFilter.Protocol
            LocalPort   = $portFilter.LocalPort
            RemotePort  = $portFilter.RemotePort
        }
    }
    $fwRules | Format-Table DisplayName, Direction, Action, Protocol, LocalPort -AutoSize
    $audit.FirewallRules = @($fwRules | ForEach-Object {
        [ordered]@{
            Name = $_.DisplayName; Direction = $_.Direction; Action = $_.Action
            Protocol = $_.Protocol; LocalPort = $_.LocalPort; RemotePort = $_.RemotePort
        }
    })
    Write-Host "  [OK] Firewall rules collected - $($fwRules.Count) enabled rules" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "FirewallRules"; Error = $_.Exception.Message }
}

# =====================================================
# 9. AD DOMAIN INFO
# =====================================================
Write-Host ""
Write-Host "=== 9. AD DOMAIN INFO ===" -ForegroundColor Cyan
Write-Host "  Running: Get-ADDomain, Get-ADForest"
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    $script:domain = Get-ADDomain
    $script:forest = Get-ADForest
    $domain = $script:domain
    $forest = $script:forest

    $adInfo = [ordered]@{
        DomainName       = $domain.DNSRoot
        NetBIOSName      = $domain.NetBIOSName
        DomainMode       = "$($domain.DomainMode)"
        ForestName       = $forest.Name
        ForestMode       = "$($forest.ForestMode)"
        PDCEmulator      = $domain.PDCEmulator
        SchemaMaster     = $forest.SchemaMaster
        Sites            = @($forest.Sites)
        GlobalCatalogs   = @($forest.GlobalCatalogs)
    }

    $adInfo.GetEnumerator() | ForEach-Object {
        $val = if ($_.Value -is [array]) { $_.Value -join ", " } else { $_.Value }
        Write-Host "  $($_.Key): $val"
    }
    $audit.ADDomain = $adInfo

    # Trusts
    Write-Host "  Running: Get-ADTrust"
    try {
        $trusts = Get-ADTrust -Filter *
        if ($trusts) {
            $trusts | Format-Table Name, Direction, TrustType, IntraForest -AutoSize
            $audit.ADTrusts = @($trusts | ForEach-Object {
                [ordered]@{ Name = $_.Name; Direction = "$($_.Direction)"; Type = "$($_.TrustType)"; IntraForest = $_.IntraForest }
            })
        } else {
            Write-Host "  No trusts configured"
            $audit.ADTrusts = @()
        }
    }
    catch {
        Write-Host "  [FAIL] Trusts: $($_.Exception.Message)" -ForegroundColor Red
        $audit._errors += @{ Section = "ADTrusts"; Error = $_.Exception.Message }
    }

    Write-Host "  [OK] AD domain info collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  (AD module not available - this may not be a domain controller)" -ForegroundColor Yellow
    $audit._errors += @{ Section = "ADDomain"; Error = $_.Exception.Message }
}

# =====================================================
# 10. DOMAIN CONTROLLERS
# =====================================================
Write-Host ""
Write-Host "=== 10. DOMAIN CONTROLLERS ===" -ForegroundColor Cyan
Write-Host "  Running: Get-ADDomainController -Filter *"
try {
    $dcs = Get-ADDomainController -Filter *
    $dcs | Format-Table HostName, IPv4Address, Site, IsGlobalCatalog, OperatingSystem -AutoSize
    $audit.DomainControllers = @($dcs | ForEach-Object {
        [ordered]@{
            Hostname = $_.HostName; IP = $_.IPv4Address; Site = $_.Site
            IsGC = $_.IsGlobalCatalog; OS = $_.OperatingSystem
        }
    })
    Write-Host "  [OK] Domain controllers collected - $($dcs.Count) DCs" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "DomainControllers"; Error = $_.Exception.Message }
}

# =====================================================
# 11. AD OU STRUCTURE
# =====================================================
Write-Host ""
Write-Host "=== 11. AD OU STRUCTURE ===" -ForegroundColor Cyan
Write-Host "  Running: Get-ADOrganizationalUnit -Filter *"
try {
    $ous = Get-ADOrganizationalUnit -Filter * -Properties Description, ProtectedFromAccidentalDeletion |
        Sort-Object DistinguishedName
    $ous | Format-Table Name, DistinguishedName, Description, ProtectedFromAccidentalDeletion -AutoSize -Wrap
    $audit.ADOUs = @($ous | ForEach-Object {
        [ordered]@{
            Name = $_.Name; DN = $_.DistinguishedName; Description = $_.Description
            Protected = $_.ProtectedFromAccidentalDeletion
        }
    })
    Write-Host "  [OK] OU structure collected - $($ous.Count) OUs" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "ADOUs"; Error = $_.Exception.Message }
}

# =====================================================
# 12. AD USERS
# =====================================================
Write-Host ""
Write-Host "=== 12. AD USERS ===" -ForegroundColor Cyan
Write-Host "  Running: Get-ADUser -Filter * -Properties ..."
try {
    $users = Get-ADUser -Filter * -Properties LastLogonDate, Enabled, Description,
        PasswordLastSet, PasswordNeverExpires, DistinguishedName, EmailAddress, WhenCreated |
        Select-Object Name, SamAccountName, Enabled, LastLogonDate, PasswordLastSet,
            PasswordNeverExpires, EmailAddress, Description, DistinguishedName, WhenCreated
    $users | Format-Table Name, SamAccountName, Enabled, LastLogonDate, EmailAddress -AutoSize
    $audit.ADUsers = @($users | ForEach-Object {
        [ordered]@{
            Name = $_.Name; SAM = $_.SamAccountName; Enabled = $_.Enabled
            LastLogon = if ($_.LastLogonDate) { $_.LastLogonDate.ToString("yyyy-MM-dd") } else { $null }
            PasswordLastSet = if ($_.PasswordLastSet) { $_.PasswordLastSet.ToString("yyyy-MM-dd") } else { $null }
            PasswordNeverExpires = $_.PasswordNeverExpires
            Email = $_.EmailAddress; Description = $_.Description
            OU = $_.DistinguishedName; Created = if ($_.WhenCreated) { $_.WhenCreated.ToString("yyyy-MM-dd") } else { $null }
        }
    })
    Write-Host "  [OK] AD users collected - $($users.Count) users" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "ADUsers"; Error = $_.Exception.Message }
}

# =====================================================
# 13. AD COMPUTERS
# =====================================================
Write-Host ""
Write-Host "=== 13. AD COMPUTERS ===" -ForegroundColor Cyan
Write-Host "  Running: Get-ADComputer -Filter * -Properties ..."
try {
    $computers = Get-ADComputer -Filter * -Properties LastLogonDate, OperatingSystem,
        OperatingSystemVersion, IPv4Address, DistinguishedName, Description, WhenCreated |
        Select-Object Name, OperatingSystem, OperatingSystemVersion, IPv4Address,
            LastLogonDate, Description, DistinguishedName, WhenCreated
    $computers | Format-Table Name, OperatingSystem, IPv4Address, LastLogonDate -AutoSize
    $audit.ADComputers = @($computers | ForEach-Object {
        [ordered]@{
            Name = $_.Name; OS = $_.OperatingSystem; OSVersion = $_.OperatingSystemVersion
            IP = $_.IPv4Address
            LastLogon = if ($_.LastLogonDate) { $_.LastLogonDate.ToString("yyyy-MM-dd") } else { $null }
            Description = $_.Description; OU = $_.DistinguishedName
            Created = if ($_.WhenCreated) { $_.WhenCreated.ToString("yyyy-MM-dd") } else { $null }
        }
    })
    Write-Host "  [OK] AD computers collected - $($computers.Count) computers" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "ADComputers"; Error = $_.Exception.Message }
}

# =====================================================
# 14. AD GROUPS WITH MEMBERS
# =====================================================
Write-Host ""
Write-Host "=== 14. AD GROUPS WITH MEMBERS ===" -ForegroundColor Cyan
Write-Host "  Running: Get-ADGroup -Filter * + Get-ADGroupMember"
try {
    $groups = Get-ADGroup -Filter * -Properties Description, Members |
        Where-Object { $_.Members.Count -gt 0 -and $_.GroupCategory -eq 'Security' } |
        Sort-Object Name

    $audit.ADGroups = @()
    foreach ($g in $groups) {
        Write-Host "  Group: $($g.Name) [$($g.GroupScope)]" -ForegroundColor Yellow
        try {
            $members = Get-ADGroupMember $g.DistinguishedName -ErrorAction Stop |
                Select-Object Name, SamAccountName, objectClass
            $members | Format-Table -AutoSize

            $audit.ADGroups += [ordered]@{
                Name = $g.Name; Scope = "$($g.GroupScope)"; Category = "$($g.GroupCategory)"
                Description = $g.Description
                Members = @($members | ForEach-Object { [ordered]@{ Name = $_.Name; SAM = $_.SamAccountName; Type = $_.objectClass } })
            }
        }
        catch {
            Write-Host "    [FAIL] Could not enumerate members: $($_.Exception.Message)" -ForegroundColor Red
            $audit.ADGroups += [ordered]@{
                Name = $g.Name; Scope = "$($g.GroupScope)"; Error = $_.Exception.Message
            }
        }
    }
    Write-Host "  [OK] AD groups collected - $($groups.Count) groups with members" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "ADGroups"; Error = $_.Exception.Message }
}

# =====================================================
# 15. PRIVILEGED GROUPS
# =====================================================
Write-Host ""
Write-Host "=== 15. PRIVILEGED GROUPS ===" -ForegroundColor Cyan
Write-Host "  Running: Get-ADGroupMember for privileged groups"

$privGroups = @("Domain Admins", "Enterprise Admins", "Schema Admins", "Administrators",
    "Account Operators", "Backup Operators", "Server Operators")
$audit.PrivilegedGroups = @()

foreach ($pg in $privGroups) {
    try {
        $members = Get-ADGroupMember $pg -ErrorAction Stop | Select-Object Name, SamAccountName, objectClass
        Write-Host "  $pg`:" -ForegroundColor Yellow
        if ($members) {
            $members | ForEach-Object { Write-Host "    $($_.Name) ($($_.SamAccountName)) [$($_.objectClass)]" }
            $audit.PrivilegedGroups += [ordered]@{
                Group = $pg
                Members = @($members | ForEach-Object { [ordered]@{ Name = $_.Name; SAM = $_.SamAccountName; Type = $_.objectClass } })
            }
        } else {
            Write-Host "    (empty)"
            $audit.PrivilegedGroups += [ordered]@{ Group = $pg; Members = @() }
        }
    }
    catch {
        Write-Host "  $pg`: [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    }
}
Write-Host "  [OK] Privileged groups checked" -ForegroundColor Green

# =====================================================
# 16. GROUP POLICY
# =====================================================
Write-Host ""
Write-Host "=== 16. GROUP POLICY ===" -ForegroundColor Cyan
Write-Host "  Running: Get-GPO -All + Get-GPOReport (for links)"
try {
    Import-Module GroupPolicy -ErrorAction Stop
    $gpos = Get-GPO -All

    $audit.GPOs = @()
    foreach ($gpo in $gpos) {
        Write-Host "  GPO: $($gpo.DisplayName)" -ForegroundColor Yellow
        Write-Host "    Status: $($gpo.GpoStatus)  Modified: $($gpo.ModificationTime)"

        $links = @()
        try {
            [xml]$report = Get-GPOReport -Guid $gpo.Id -ReportType Xml -ErrorAction Stop
            if ($report.GPO.LinksTo) {
                $report.GPO.LinksTo | ForEach-Object {
                    Write-Host "    Link: $($_.SOMPath) [Enabled: $($_.Enabled)]"
                    $links += [ordered]@{ Path = $_.SOMPath; Enabled = $_.Enabled }
                }
            } else {
                Write-Host "    Link: (not linked)"
            }
        }
        catch {
            Write-Host "    [FAIL] Could not get links: $($_.Exception.Message)" -ForegroundColor Red
        }

        $audit.GPOs += [ordered]@{
            Name = $gpo.DisplayName; Status = "$($gpo.GpoStatus)"
            Created = $gpo.CreationTime.ToString("yyyy-MM-dd")
            Modified = $gpo.ModificationTime.ToString("yyyy-MM-dd")
            Links = $links
        }
    }
    Write-Host "  [OK] GPOs collected - $($gpos.Count) GPOs" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "GPOs"; Error = $_.Exception.Message }
}

# =====================================================
# 17. LOCAL USERS
# =====================================================
Write-Host ""
Write-Host "=== 17. LOCAL USERS ===" -ForegroundColor Cyan
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
# 18. LOCAL ADMINISTRATORS
# =====================================================
Write-Host ""
Write-Host "=== 18. LOCAL ADMINISTRATORS ===" -ForegroundColor Cyan
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
# 19. SHARES WITH PERMISSIONS
# =====================================================
Write-Host ""
Write-Host "=== 19. SHARES WITH PERMISSIONS ===" -ForegroundColor Cyan
Write-Host "  Running: Get-SmbShare + Get-SmbShareAccess + Get-Acl"
try {
    $shares = Get-SmbShare | Where-Object { $_.Name -notlike '*$' }
    $audit.Shares = @()

    foreach ($s in $shares) {
        Write-Host "  Share: \\$env:COMPUTERNAME\$($s.Name)" -ForegroundColor Yellow
        Write-Host "    Path: $($s.Path)"
        Write-Host "    Description: $($s.Description)"

        $shareData = [ordered]@{
            Name = $s.Name; Path = $s.Path; Description = $s.Description
            SMBPermissions = @(); NTFSPermissions = @()
        }

        # SMB permissions
        Write-Host "    SMB Permissions:" -ForegroundColor DarkGray
        try {
            $smbPerms = Get-SmbShareAccess -Name $s.Name
            $smbPerms | ForEach-Object {
                Write-Host "      $($_.AccountName): $($_.AccessRight) ($($_.AccessControlType))"
                $shareData.SMBPermissions += [ordered]@{
                    Account = $_.AccountName; Right = "$($_.AccessRight)"; Type = "$($_.AccessControlType)"
                }
            }
        }
        catch {
            Write-Host "      [FAIL] $($_.Exception.Message)" -ForegroundColor Red
        }

        # NTFS permissions
        Write-Host "    NTFS Permissions:" -ForegroundColor DarkGray
        try {
            if (Test-Path $s.Path) {
                $acl = Get-Acl $s.Path
                $acl.Access | ForEach-Object {
                    Write-Host "      $($_.IdentityReference): $($_.FileSystemRights) ($($_.AccessControlType))"
                    $shareData.NTFSPermissions += [ordered]@{
                        Identity = "$($_.IdentityReference)"; Rights = "$($_.FileSystemRights)"
                        Type = "$($_.AccessControlType)"; Inherited = $_.IsInherited
                    }
                }
            }
        }
        catch {
            Write-Host "      [FAIL] $($_.Exception.Message)" -ForegroundColor Red
        }

        $audit.Shares += $shareData
    }
    Write-Host "  [OK] Shares collected - $($shares.Count) shares" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "Shares"; Error = $_.Exception.Message }
}

# =====================================================
# 20. INSTALLED ROLES/FEATURES
# =====================================================
Write-Host ""
Write-Host "=== 20. INSTALLED ROLES/FEATURES ===" -ForegroundColor Cyan
Write-Host "  Running: Get-WindowsFeature | Where Installed"
try {
    $roles = Get-WindowsFeature | Where-Object { $_.Installed } | Select-Object Name, DisplayName
    $roles | Format-Table -AutoSize
    $audit.InstalledRoles = @($roles | ForEach-Object {
        [ordered]@{ Name = $_.Name; DisplayName = $_.DisplayName }
    })
    Write-Host "  [OK] Roles collected - $($roles.Count) installed" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "InstalledRoles"; Error = $_.Exception.Message }
}

# =====================================================
# 21. DHCP
# =====================================================
Write-Host ""
Write-Host "=== 21. DHCP ===" -ForegroundColor Cyan
Write-Host "  Running: Get-DhcpServerv4Scope + options + reservations + leases"
try {
    Import-Module DhcpServer -ErrorAction Stop
    $scopes = Get-DhcpServerv4Scope
    $audit.DHCP = [ordered]@{ Scopes = @() }

    # Server-level options
    Write-Host "  Server-level DHCP options:" -ForegroundColor Yellow
    try {
        $serverOpts = Get-DhcpServerv4OptionValue -ErrorAction Stop
        $serverOpts | Format-Table OptionId, Name, Value -AutoSize
        $audit.DHCP.ServerOptions = @($serverOpts | ForEach-Object {
            [ordered]@{ OptionId = $_.OptionId; Name = $_.Name; Value = @($_.Value) }
        })
    }
    catch {
        Write-Host "    [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    }

    # Failover
    try {
        $failover = Get-DhcpServerv4Failover -ErrorAction Stop
        if ($failover) {
            $failover | Format-Table Name, PartnerServer, Mode, State -AutoSize
            $audit.DHCP.Failover = @($failover | ForEach-Object {
                [ordered]@{ Name = $_.Name; Partner = $_.PartnerServer; Mode = "$($_.Mode)"; State = "$($_.State)" }
            })
        }
    }
    catch { Write-Host "  No DHCP failover configured" -ForegroundColor DarkGray }

    foreach ($scope in $scopes) {
        $scopeData = [ordered]@{
            ScopeId = "$($scope.ScopeId)"; Name = $scope.Name
            SubnetMask = "$($scope.SubnetMask)"; StartRange = "$($scope.StartRange)"
            EndRange = "$($scope.EndRange)"; State = "$($scope.State)"
            LeaseDuration = "$($scope.LeaseDuration)"
            Options = @(); Exclusions = @(); Reservations = @(); Leases = @()
        }

        Write-Host "  Scope: $($scope.ScopeId) ($($scope.Name))" -ForegroundColor Yellow
        Write-Host "    Range: $($scope.StartRange) - $($scope.EndRange), Mask: $($scope.SubnetMask), State: $($scope.State)"

        # Scope options
        try {
            $scopeOpts = Get-DhcpServerv4OptionValue -ScopeId $scope.ScopeId -ErrorAction Stop
            Write-Host "    Options:"
            $scopeOpts | ForEach-Object { Write-Host "      $($_.OptionId) $($_.Name): $($_.Value -join ', ')" }
            $scopeData.Options = @($scopeOpts | ForEach-Object {
                [ordered]@{ OptionId = $_.OptionId; Name = $_.Name; Value = @($_.Value) }
            })
        }
        catch { Write-Host "    Options: [FAIL] $($_.Exception.Message)" -ForegroundColor Red }

        # Exclusions
        try {
            $exclusions = Get-DhcpServerv4ExclusionRange -ScopeId $scope.ScopeId -ErrorAction Stop
            if ($exclusions) {
                Write-Host "    Exclusions:"
                $exclusions | ForEach-Object { Write-Host "      $($_.StartRange) - $($_.EndRange)" }
                $scopeData.Exclusions = @($exclusions | ForEach-Object {
                    [ordered]@{ Start = "$($_.StartRange)"; End = "$($_.EndRange)" }
                })
            }
        }
        catch {}

        # Reservations
        try {
            $reservations = Get-DhcpServerv4Reservation -ScopeId $scope.ScopeId -ErrorAction Stop
            if ($reservations) {
                Write-Host "    Reservations:"
                $reservations | Format-Table IPAddress, Name, ClientId, Description -AutoSize
                $scopeData.Reservations = @($reservations | ForEach-Object {
                    [ordered]@{ IP = "$($_.IPAddress)"; Name = $_.Name; MAC = $_.ClientId; Description = $_.Description }
                })
            }
        }
        catch { Write-Host "    Reservations: [FAIL] $($_.Exception.Message)" -ForegroundColor Red }

        # Leases
        try {
            $leases = Get-DhcpServerv4Lease -ScopeId $scope.ScopeId -ErrorAction Stop
            if ($leases) {
                Write-Host "    Active Leases:"
                $leases | Format-Table IPAddress, HostName, ClientId, AddressState -AutoSize
                $scopeData.Leases = @($leases | ForEach-Object {
                    [ordered]@{ IP = "$($_.IPAddress)"; Hostname = $_.HostName; MAC = $_.ClientId; State = "$($_.AddressState)" }
                })
            }
        }
        catch { Write-Host "    Leases: [FAIL] $($_.Exception.Message)" -ForegroundColor Red }

        $audit.DHCP.Scopes += $scopeData
    }
    Write-Host "  [OK] DHCP collected - $($scopes.Count) scopes" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  (DHCP role not installed or not running)" -ForegroundColor Yellow
    $audit._errors += @{ Section = "DHCP"; Error = $_.Exception.Message }
}

# =====================================================
# 22. DNS SERVER
# =====================================================
Write-Host ""
Write-Host "=== 22. DNS SERVER ===" -ForegroundColor Cyan
Write-Host "  Running: Get-DnsServerZone, Get-DnsServerForwarder, Get-DnsServerResourceRecord"
try {
    Import-Module DnsServer -ErrorAction Stop

    # Forwarders
    Write-Host "  Forwarders:" -ForegroundColor Yellow
    $fwd = Get-DnsServerForwarder
    $fwd.IPAddress | ForEach-Object { Write-Host "    $_" }
    $audit.DNS = [ordered]@{
        Forwarders = @($fwd.IPAddress | ForEach-Object { "$_" })
        ConditionalForwarders = @()
        Zones = @()
    }

    # Conditional forwarders
    Write-Host "  Conditional Forwarders:" -ForegroundColor Yellow
    $condFwd = Get-DnsServerZone | Where-Object { $_.ZoneType -eq 'Forwarder' }
    if ($condFwd) {
        $condFwd | ForEach-Object {
            Write-Host "    $($_.ZoneName) -> $($_.MasterServers -join ', ')"
            $audit.DNS.ConditionalForwarders += [ordered]@{ Zone = $_.ZoneName; ForwardTo = @($_.MasterServers | ForEach-Object { "$_" }) }
        }
    } else {
        Write-Host "    (none)"
    }

    # Zones and records
    Write-Host "  Zones:" -ForegroundColor Yellow
    $zones = Get-DnsServerZone | Where-Object { $_.ZoneType -ne 'Forwarder' }
    foreach ($z in $zones) {
        $zoneData = [ordered]@{
            Name = $z.ZoneName; Type = "$($z.ZoneType)"; IsReverse = $z.IsReverseLookupZone
            DynamicUpdate = "$($z.DynamicUpdate)"; Records = @()
        }

        Write-Host "    $($z.ZoneName) [$($z.ZoneType)] Reverse=$($z.IsReverseLookupZone)" -ForegroundColor DarkGray

        # Get records for primary zones (skip if >500 records to avoid flooding)
        if ($z.ZoneType -eq 'Primary' -and -not $z.IsAutoCreated) {
            try {
                $records = Get-DnsServerResourceRecord -ZoneName $z.ZoneName -ErrorAction Stop
                if ($records.Count -le 500) {
                    $zoneData.Records = @($records | ForEach-Object {
                        [ordered]@{
                            Name = $_.HostName; Type = "$($_.RecordType)"; TTL = "$($_.TimeToLive)"
                            Data = "$($_.RecordData.IPv4Address)$($_.RecordData.HostNameAlias)$($_.RecordData.NameServer)$($_.RecordData.DescriptiveText)$($_.RecordData.DomainName)"
                        }
                    })
                    Write-Host "      $($records.Count) records"
                } else {
                    Write-Host "      $($records.Count) records (too many to list, count only)" -ForegroundColor Yellow
                    $zoneData.RecordCount = $records.Count
                }
            }
            catch {
                Write-Host "      [FAIL] Records: $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        $audit.DNS.Zones += $zoneData
    }
    Write-Host "  [OK] DNS collected - $($zones.Count) zones" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  (DNS Server role not installed)" -ForegroundColor Yellow
    $audit._errors += @{ Section = "DNS"; Error = $_.Exception.Message }
}

# =====================================================
# 23. HYPER-V VMs
# =====================================================
Write-Host ""
Write-Host "=== 23. HYPER-V VMs ===" -ForegroundColor Cyan
Write-Host "  Running: Get-VM"
try {
    $vms = Get-VM -ErrorAction Stop
    if ($vms) {
        $audit.HyperV = @($vms | ForEach-Object {
            $vm = $_
            Write-Host "  VM: $($vm.Name) [State: $($vm.State)]" -ForegroundColor Yellow
            Write-Host "    CPU: $($vm.ProcessorCount), Memory: $([math]::Round($vm.MemoryAssigned/1MB))MB, Gen: $($vm.Generation)"

            $netAdapters = Get-VMNetworkAdapter -VMName $vm.Name -ErrorAction SilentlyContinue
            $disks = Get-VMHardDiskDrive -VMName $vm.Name -ErrorAction SilentlyContinue

            [ordered]@{
                Name = $vm.Name; State = "$($vm.State)"; CPUs = $vm.ProcessorCount
                MemoryMB = [math]::Round($vm.MemoryAssigned/1MB); Generation = $vm.Generation
                Networks = @($netAdapters | ForEach-Object { [ordered]@{ Switch = $_.SwitchName; MAC = $_.MacAddress; IPs = @($_.IPAddresses) } })
                Disks = @($disks | ForEach-Object { $_.Path })
            }
        })
        Write-Host "  [OK] Hyper-V VMs collected - $($vms.Count) VMs" -ForegroundColor Green
    } else {
        Write-Host "  No VMs found" -ForegroundColor DarkGray
        $audit.HyperV = @()
    }
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  (Hyper-V role not installed)" -ForegroundColor Yellow
    $audit._errors += @{ Section = "HyperV"; Error = $_.Exception.Message }
}

# =====================================================
# 24. IIS SITES
# =====================================================
Write-Host ""
Write-Host "=== 24. IIS SITES ===" -ForegroundColor Cyan
Write-Host "  Running: Get-Website"
try {
    Import-Module WebAdministration -ErrorAction Stop
    $sites = Get-Website
    if ($sites) {
        $audit.IIS = @($sites | ForEach-Object {
            Write-Host "  Site: $($_.Name) [State: $($_.State)]" -ForegroundColor Yellow
            Write-Host "    Path: $($_.PhysicalPath)"
            Write-Host "    Bindings: $(($_.Bindings.Collection | ForEach-Object { $_.bindingInformation }) -join ', ')"
            [ordered]@{
                Name = $_.Name; State = "$($_.State)"; Path = $_.PhysicalPath
                AppPool = $_.applicationPool
                Bindings = @($_.Bindings.Collection | ForEach-Object { $_.bindingInformation })
            }
        })
        Write-Host "  [OK] IIS sites collected" -ForegroundColor Green
    } else {
        Write-Host "  No websites configured" -ForegroundColor DarkGray
    }
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  (IIS not installed)" -ForegroundColor Yellow
    $audit._errors += @{ Section = "IIS"; Error = $_.Exception.Message }
}

# =====================================================
# 25. NPS / RADIUS
# =====================================================
Write-Host ""
Write-Host "=== 25. NPS / RADIUS ===" -ForegroundColor Cyan
Write-Host "  Running: Get-NpsRadiusClient, Get-NpsNetworkPolicy"
try {
    Import-Module NPS -ErrorAction Stop
    $clients = Get-NpsRadiusClient -ErrorAction Stop
    $policies = Get-NpsNetworkPolicy -ErrorAction Stop

    $audit.NPS = [ordered]@{
        Clients = @($clients | ForEach-Object {
            Write-Host "  RADIUS Client: $($_.Name) ($($_.Address))" -ForegroundColor Yellow
            [ordered]@{ Name = $_.Name; Address = $_.Address }
        })
        Policies = @($policies | ForEach-Object {
            Write-Host "  Policy: $($_.Name) [Enabled: $($_.Enabled)]" -ForegroundColor Yellow
            [ordered]@{ Name = $_.Name; Enabled = $_.Enabled; Order = $_.ProcessingOrder }
        })
    }
    Write-Host "  [OK] NPS collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  (NPS/RADIUS not installed)" -ForegroundColor Yellow
    $audit._errors += @{ Section = "NPS"; Error = $_.Exception.Message }
}

# =====================================================
# 26. WINDOWS SERVER BACKUP
# =====================================================
Write-Host ""
Write-Host "=== 26. WINDOWS SERVER BACKUP ===" -ForegroundColor Cyan
Write-Host "  Running: Get-WBPolicy, Get-WBSummary"
try {
    Add-PSSnapin Windows.ServerBackup -ErrorAction Stop
    $policy = Get-WBPolicy -ErrorAction Stop

    $wsbData = [ordered]@{
        Schedule = @(Get-WBSchedule -Policy $policy)
        Volumes = @((Get-WBVolume -Policy $policy).MountPoint)
        SystemState = (Get-WBSystemState -Policy $policy)
    }

    try {
        $target = Get-WBBackupTarget -Policy $policy
        $wsbData.Target = "$($target.Label) ($($target.TargetPath))"
    }
    catch { $wsbData.Target = "Unknown" }

    try {
        $summary = Get-WBSummary
        $wsbData.LastSuccess = if ($summary.LastSuccessfulBackupTime) { $summary.LastSuccessfulBackupTime.ToString("yyyy-MM-dd HH:mm") } else { $null }
        $wsbData.LastResult = "$($summary.LastBackupResultHR)"
        $wsbData.NextRun = if ($summary.NextBackupTime) { $summary.NextBackupTime.ToString("yyyy-MM-dd HH:mm") } else { $null }
    }
    catch {}

    $wsbData.GetEnumerator() | ForEach-Object {
        $val = if ($_.Value -is [array]) { $_.Value -join ", " } else { $_.Value }
        Write-Host "  $($_.Key): $val"
    }

    $audit.WindowsServerBackup = $wsbData
    Write-Host "  [OK] WSB info collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  (Windows Server Backup not installed or no policy configured)" -ForegroundColor Yellow
    $audit._errors += @{ Section = "WindowsServerBackup"; Error = $_.Exception.Message }
}

# =====================================================
# 27. PRINTERS
# =====================================================
Write-Host ""
Write-Host "=== 27. PRINTERS ===" -ForegroundColor Cyan
Write-Host "  Running: Get-Printer, Get-PrinterPort"
try {
    $printers = Get-Printer | Select-Object Name, DriverName, PortName, Shared, ShareName, Published
    $printers | Format-Table -AutoSize

    $ports = Get-PrinterPort | Where-Object { $_.Name -like 'TCP_*' -or $_.Name -like 'IP_*' } |
        Select-Object Name, PrinterHostAddress, PortNumber
    if ($ports) {
        Write-Host "  Printer Ports:" -ForegroundColor DarkGray
        $ports | Format-Table -AutoSize
    }

    $audit.Printers = @($printers | ForEach-Object {
        [ordered]@{
            Name = $_.Name; Driver = $_.DriverName; Port = $_.PortName
            Shared = $_.Shared; ShareName = $_.ShareName
        }
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
# 28. CERTIFICATES
# =====================================================
Write-Host ""
Write-Host "=== 28. CERTIFICATES ===" -ForegroundColor Cyan
Write-Host "  Running: Get-ChildItem Cert:\LocalMachine\My"
try {
    $certs = Get-ChildItem Cert:\LocalMachine\My | Select-Object Subject, NotAfter, NotBefore, Issuer, Thumbprint
    $certs | Format-Table Subject, NotAfter, Issuer -AutoSize

    $expiringSoon = $certs | Where-Object { $_.NotAfter -lt (Get-Date).AddDays(90) }
    if ($expiringSoon) {
        Write-Host "  WARNING: Certificates expiring within 90 days:" -ForegroundColor Yellow
        $expiringSoon | ForEach-Object { Write-Host "    $($_.Subject) expires $($_.NotAfter)" -ForegroundColor Yellow }
    }

    $audit.Certificates = @($certs | ForEach-Object {
        [ordered]@{
            Subject = $_.Subject; Issuer = $_.Issuer; Thumbprint = $_.Thumbprint
            NotBefore = $_.NotBefore.ToString("yyyy-MM-dd"); NotAfter = $_.NotAfter.ToString("yyyy-MM-dd")
            ExpiringSoon = ($_.NotAfter -lt (Get-Date).AddDays(90))
        }
    })
    Write-Host "  [OK] Certificates collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "Certificates"; Error = $_.Exception.Message }
}

# =====================================================
# 29. INSTALLED SOFTWARE (registry-based)
# =====================================================
Write-Host ""
Write-Host "=== 29. INSTALLED SOFTWARE ===" -ForegroundColor Cyan
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
        [ordered]@{
            Name = $_.DisplayName; Version = $_.DisplayVersion
            Publisher = $_.Publisher; InstallDate = $_.InstallDate
        }
    })
    Write-Host "  [OK] Software collected - $($software.Count) packages" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "InstalledSoftware"; Error = $_.Exception.Message }
}

# =====================================================
# 30. SERVICES (non-default running)
# =====================================================
Write-Host ""
Write-Host "=== 30. SERVICES (non-default, running) ===" -ForegroundColor Cyan
Write-Host "  Running: Get-Service | Where Running + Automatic"
try {
    $services = Get-Service | Where-Object { $_.Status -eq 'Running' -and $_.StartType -eq 'Automatic' } |
        Where-Object { $_.DisplayName -notmatch '^(Windows|Microsoft|Net\.|COM\+|DCOM|WMI|Plug|Task|CNG|Base|Crypto|Security Center|DHCP Client|DNS Client|Group Policy|IP Helper|Server$|Workstation$|TCP/IP|Remote Procedure|User Profile|Background|Connected|CoreMessaging|State Repository|Storage|System Events|Time Broker|User Manager|WinHTTP|Diagnostic)' } |
        Select-Object Name, DisplayName, StartType |
        Sort-Object DisplayName

    $services | Format-Table -AutoSize

    $audit.Services = @($services | ForEach-Object {
        [ordered]@{ Name = $_.Name; DisplayName = $_.DisplayName; StartType = "$($_.StartType)" }
    })
    Write-Host "  [OK] Services collected - $($services.Count) non-default" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "Services"; Error = $_.Exception.Message }
}

# =====================================================
# 31. SCHEDULED TASKS (non-Microsoft)
# =====================================================
Write-Host ""
Write-Host "=== 31. SCHEDULED TASKS (non-Microsoft) ===" -ForegroundColor Cyan
Write-Host "  Running: Get-ScheduledTask"
try {
    $tasks = Get-ScheduledTask |
        Where-Object { $_.Author -notlike 'Microsoft*' -and $_.State -ne 'Disabled' } |
        Select-Object TaskName, State, Author, TaskPath
    $tasks | Format-Table -AutoSize

    $audit.ScheduledTasks = @($tasks | ForEach-Object {
        [ordered]@{ Name = $_.TaskName; State = "$($_.State)"; Author = $_.Author; Path = $_.TaskPath }
    })
    Write-Host "  [OK] Scheduled tasks collected - $($tasks.Count) non-Microsoft" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "ScheduledTasks"; Error = $_.Exception.Message }
}

# =====================================================
# 32. WINDOWS UPDATES
# =====================================================
Write-Host ""
Write-Host "=== 32. WINDOWS UPDATES (last 20) ===" -ForegroundColor Cyan
Write-Host "  Running: Get-HotFix"
try {
    $updates = Get-HotFix | Sort-Object InstalledOn -Descending -ErrorAction SilentlyContinue | Select-Object -First 20 |
        Select-Object HotFixID, Description, InstalledOn, InstalledBy
    $updates | Format-Table -AutoSize

    $audit.WindowsUpdates = @($updates | ForEach-Object {
        [ordered]@{
            KB = $_.HotFixID; Description = $_.Description
            InstalledOn = if ($_.InstalledOn) { $_.InstalledOn.ToString("yyyy-MM-dd") } else { $null }
            InstalledBy = $_.InstalledBy
        }
    })
    Write-Host "  [OK] Updates collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "WindowsUpdates"; Error = $_.Exception.Message }
}

# =====================================================
# 33. TIME CONFIG
# =====================================================
Write-Host ""
Write-Host "=== 33. TIME CONFIG ===" -ForegroundColor Cyan
Write-Host "  Running: w32tm /query /configuration"
try {
    $timeConfig = w32tm /query /configuration 2>&1
    $timeConfig | ForEach-Object { Write-Host "  $_" }
    Write-Host "  [OK] Time config collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "TimeConfig"; Error = $_.Exception.Message }
}

# =====================================================
# 34. EVENT LOG ERRORS (last 14 days)
# =====================================================
Write-Host ""
Write-Host "=== 34. EVENT LOG ERRORS (Critical/Error, last 14 days) ===" -ForegroundColor Cyan
Write-Host "  Running: Get-WinEvent -FilterHashtable System log"
try {
    $since = (Get-Date).AddDays(-14)
    $events = Get-WinEvent -FilterHashtable @{LogName='System'; Level=1,2; StartTime=$since} -MaxEvents 30 -ErrorAction Stop |
        Select-Object TimeCreated, Id, LevelDisplayName, ProviderName, @{N='Message';E={$_.Message.Substring(0, [Math]::Min(200, $_.Message.Length))}}
    $events | Format-Table TimeCreated, Id, ProviderName -AutoSize

    $audit.EventLogErrors = @($events | ForEach-Object {
        [ordered]@{
            Time = $_.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
            EventId = $_.Id; Level = $_.LevelDisplayName; Source = $_.ProviderName
            Message = $_.Message
        }
    })
    Write-Host "  [OK] Events collected - $($events.Count) critical/error events" -ForegroundColor Green
}
catch {
    Write-Host "  No critical/error events in System log (last 14 days)" -ForegroundColor Green
    $audit.EventLogErrors = @()
}

# =====================================================
# 35. EVENT LOG WARNINGS (last 14 days)
# =====================================================
Write-Host ""
Write-Host "=== 35. EVENT LOG WARNINGS (last 14 days) ===" -ForegroundColor Cyan
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
# 36. EVENT LOG SETTINGS
# =====================================================
Write-Host ""
Write-Host "=== 36. EVENT LOG SETTINGS ===" -ForegroundColor Cyan
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
            Write-Host "  $($log.LogName): Max=$($logInfo.MaxSizeKB)KB, Current=$($logInfo.CurrentSizeKB)KB, Mode=$($log.LogMode), Records=$($log.RecordCount)"
            $audit.EventLogSettings += $logInfo
        }
        catch { Write-Host "  $logName`: not available" -ForegroundColor DarkGray }
    }

    # HIPAA requires adequate log retention
    $securityLog = $audit.EventLogSettings | Where-Object { $_.LogName -eq 'Security' }
    if ($securityLog -and $securityLog.MaxSizeKB -lt 131072) {
        Write-Host "  [WARN] Security log max size is $($securityLog.MaxSizeKB)KB - recommend at least 128MB for compliance" -ForegroundColor Yellow
    }
    if ($securityLog -and $securityLog.LogMode -eq 'Circular') {
        Write-Host "  [WARN] Security log is in Circular mode - old events will be overwritten. Consider archiving for HIPAA" -ForegroundColor Yellow
    }
    Write-Host "  [OK] Event log settings collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "EventLogSettings"; Error = $_.Exception.Message }
}

# =====================================================
# 37. WINDOWS FIREWALL PROFILES
# =====================================================
Write-Host ""
Write-Host "=== 37. WINDOWS FIREWALL PROFILES ===" -ForegroundColor Cyan
Write-Host "  Running: Get-NetFirewallProfile"
try {
    $fwProfiles = Get-NetFirewallProfile -ErrorAction Stop
    $audit.FirewallProfiles = @($fwProfiles | ForEach-Object {
        $p = [ordered]@{
            Profile = "$($_.Name)"; Enabled = $_.Enabled
            DefaultInboundAction = "$($_.DefaultInboundAction)"
            DefaultOutboundAction = "$($_.DefaultOutboundAction)"
            LogAllowed = $_.LogAllowed; LogBlocked = $_.LogBlocked
            LogFileName = $_.LogFileName; LogMaxSizeKilobytes = $_.LogMaxSizeKilobytes
        }
        Write-Host "  $($_.Name): Enabled=$($_.Enabled), Inbound=$($_.DefaultInboundAction), Outbound=$($_.DefaultOutboundAction)"
        if (-not $_.Enabled) {
            Write-Host "    [WARN] $($_.Name) firewall profile is DISABLED" -ForegroundColor Red
        }
        if (-not $_.LogBlocked) {
            Write-Host "    [WARN] $($_.Name) blocked connections are NOT being logged" -ForegroundColor Yellow
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
# 38. RDP SECURITY SETTINGS
# =====================================================
Write-Host ""
Write-Host "=== 38. RDP SECURITY SETTINGS ===" -ForegroundColor Cyan
Write-Host "  Running: Registry check for RDP settings"
try {
    $tsReg = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server'
    $rdpReg = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp'

    $rdpEnabled = (Get-ItemProperty $tsReg -ErrorAction Stop).fDenyTSConnections
    $nla = (Get-ItemProperty $rdpReg -ErrorAction SilentlyContinue).UserAuthentication
    $secLayer = (Get-ItemProperty $rdpReg -ErrorAction SilentlyContinue).SecurityLayer
    $minEncrypt = (Get-ItemProperty $rdpReg -ErrorAction SilentlyContinue).MinEncryptionLevel

    $rdpSettings = [ordered]@{
        RDPEnabled = ($rdpEnabled -eq 0)
        NLARequired = ($nla -eq 1)
        SecurityLayer = switch ($secLayer) { 0 {"RDP Security"} 1 {"Negotiate"} 2 {"SSL/TLS"} default {"Unknown ($secLayer)"} }
        MinEncryptionLevel = switch ($minEncrypt) { 1 {"Low"} 2 {"Client Compatible"} 3 {"High"} 4 {"FIPS"} default {"Unknown ($minEncrypt)"} }
    }
    $rdpSettings.GetEnumerator() | ForEach-Object { Write-Host "  $($_.Key): $($_.Value)" }

    if ($rdpSettings.RDPEnabled -and -not $rdpSettings.NLARequired) {
        Write-Host "  [WARN] RDP is enabled but NLA is NOT required - credential theft risk" -ForegroundColor Red
    }
    if ($rdpSettings.RDPEnabled -and $secLayer -lt 2) {
        Write-Host "  [WARN] RDP security layer is not SSL/TLS - downgrade attack risk" -ForegroundColor Yellow
    }

    $audit.RDPSettings = $rdpSettings
    Write-Host "  [OK] RDP settings collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "RDPSettings"; Error = $_.Exception.Message }
}

# =====================================================
# 39. TLS/SSL CONFIGURATION
# =====================================================
Write-Host ""
Write-Host "=== 39. TLS/SSL CONFIGURATION ===" -ForegroundColor Cyan
Write-Host "  Running: Registry check for TLS protocol versions"
try {
    $tlsVersions = @('SSL 2.0', 'SSL 3.0', 'TLS 1.0', 'TLS 1.1', 'TLS 1.2', 'TLS 1.3')
    $audit.TLSConfig = [ordered]@{}

    foreach ($ver in $tlsVersions) {
        $serverPath = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$ver\Server"
        $enabled = $null
        if (Test-Path $serverPath) {
            $enabled = (Get-ItemProperty $serverPath -ErrorAction SilentlyContinue).Enabled
        }
        $status = if ($null -eq $enabled) { "OS Default" } elseif ($enabled -eq 0) { "Disabled" } else { "Enabled" }
        Write-Host "  $ver`: $status"
        $audit.TLSConfig[$ver] = $status

        if ($ver -in @('SSL 2.0', 'SSL 3.0', 'TLS 1.0', 'TLS 1.1') -and $status -ne 'Disabled') {
            Write-Host "    [WARN] $ver should be explicitly DISABLED for compliance" -ForegroundColor Yellow
        }
    }
    Write-Host "  [OK] TLS config collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "TLSConfig"; Error = $_.Exception.Message }
}

# =====================================================
# 40. CREDENTIAL GUARD / DEVICE GUARD
# =====================================================
Write-Host ""
Write-Host "=== 40. CREDENTIAL GUARD / DEVICE GUARD ===" -ForegroundColor Cyan
Write-Host "  Running: Get-CimInstance Win32_DeviceGuard"
try {
    $dg = Get-CimInstance -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard -ErrorAction Stop
    $cgStatus = switch ($dg.SecurityServicesRunning) {
        { 1 -in $_ } { "Running" }
        default { "Not running" }
    }
    $vbsStatus = switch ($dg.VirtualizationBasedSecurityStatus) {
        0 { "Not enabled" } 1 { "Enabled but not running" } 2 { "Running" } default { "Unknown" }
    }

    $guardInfo = [ordered]@{
        VBSStatus = $vbsStatus
        CredentialGuard = $cgStatus
        SecurityServicesConfigured = @($dg.SecurityServicesConfigured)
        SecurityServicesRunning = @($dg.SecurityServicesRunning)
    }
    $guardInfo.GetEnumerator() | ForEach-Object {
        $val = if ($_.Value -is [array]) { $_.Value -join ", " } else { $_.Value }
        Write-Host "  $($_.Key): $val"
    }
    if ($cgStatus -ne "Running") {
        Write-Host "  [WARN] Credential Guard is not running - pass-the-hash risk" -ForegroundColor Yellow
    }
    $audit.CredentialGuard = $guardInfo
    Write-Host "  [OK] Credential/Device Guard checked" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "CredentialGuard"; Error = $_.Exception.Message }
}

# =====================================================
# 41. UAC SETTINGS
# =====================================================
Write-Host ""
Write-Host "=== 41. UAC SETTINGS ===" -ForegroundColor Cyan
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
        PromptOnSecureDesktop = $uacReg.PromptOnSecureDesktop
        FilterAdministratorToken = $uacReg.FilterAdministratorToken
    }
    $uacSettings.GetEnumerator() | ForEach-Object { Write-Host "  $($_.Key): $($_.Value)" }

    if ($uacReg.EnableLUA -ne 1) {
        Write-Host "  [WARN] UAC is DISABLED - all processes run with full admin rights" -ForegroundColor Red
    }
    if ($uacReg.ConsentPromptBehaviorAdmin -eq 0) {
        Write-Host "  [WARN] Admin elevation without prompting - malware risk" -ForegroundColor Yellow
    }
    $audit.UACSettings = $uacSettings
    Write-Host "  [OK] UAC settings collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "UACSettings"; Error = $_.Exception.Message }
}

# =====================================================
# 42. SCREEN LOCK / INACTIVITY TIMEOUT
# =====================================================
Write-Host ""
Write-Host "=== 42. SCREEN LOCK / INACTIVITY TIMEOUT ===" -ForegroundColor Cyan
Write-Host "  Running: Registry check for screen lock policy"
try {
    $screenLock = [ordered]@{}

    # Machine inactivity limit
    $inactivity = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -ErrorAction SilentlyContinue).InactivityTimeoutSecs
    $screenLock.InactivityTimeoutSecs = $inactivity
    if ($inactivity) {
        Write-Host "  Machine inactivity timeout: $inactivity seconds ($([math]::Round($inactivity/60)) minutes)"
    } else {
        Write-Host "  Machine inactivity timeout: Not configured"
        Write-Host "  [WARN] No machine inactivity timeout - HIPAA requires automatic session lock" -ForegroundColor Yellow
    }

    # Screensaver settings (applied via GPO)
    $ssActive = (Get-ItemProperty 'HKCU:\Control Panel\Desktop' -ErrorAction SilentlyContinue).ScreenSaveActive
    $ssTimeout = (Get-ItemProperty 'HKCU:\Control Panel\Desktop' -ErrorAction SilentlyContinue).ScreenSaveTimeOut
    $ssSecure = (Get-ItemProperty 'HKCU:\Control Panel\Desktop' -ErrorAction SilentlyContinue).ScreenSaverIsSecure
    $screenLock.ScreenSaverActive = $ssActive
    $screenLock.ScreenSaverTimeout = $ssTimeout
    $screenLock.ScreenSaverSecure = $ssSecure

    Write-Host "  Screensaver active: $ssActive, Timeout: $ssTimeout sec, Require password: $ssSecure"

    $audit.ScreenLockPolicy = $screenLock
    Write-Host "  [OK] Screen lock settings collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "ScreenLock"; Error = $_.Exception.Message }
}

# =====================================================
# 43. USB STORAGE POLICY
# =====================================================
Write-Host ""
Write-Host "=== 43. USB STORAGE POLICY ===" -ForegroundColor Cyan
Write-Host "  Running: Registry check for USB storage restrictions"
try {
    $usbStorage = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Services\USBSTOR' -ErrorAction Stop).Start
    $usbStatus = switch ($usbStorage) {
        3 { "Enabled (normal)" } 4 { "Disabled" } default { "Unknown ($usbStorage)" }
    }
    Write-Host "  USB Storage: $usbStatus"

    # Check for GPO-based USB restrictions
    $usbGPO = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\RemovableStorageDevices' -ErrorAction SilentlyContinue
    if ($usbGPO) {
        Write-Host "  GPO removable storage restrictions detected"
    }

    if ($usbStorage -eq 3 -and -not $usbGPO) {
        Write-Host "  [WARN] USB storage is unrestricted - data exfiltration risk / HIPAA concern" -ForegroundColor Yellow
    }

    $audit.USBStoragePolicy = [ordered]@{ USBSTORStart = $usbStorage; Status = $usbStatus; GPORestrictions = ($null -ne $usbGPO) }
    Write-Host "  [OK] USB storage policy collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "USBStorage"; Error = $_.Exception.Message }
}

# =====================================================
# 44. FAILED LOGON ATTEMPTS (last 7 days)
# =====================================================
Write-Host ""
Write-Host "=== 44. FAILED LOGON ATTEMPTS (last 7 days) ===" -ForegroundColor Cyan
Write-Host "  Running: Get-WinEvent Security log - Event ID 4625"
$failedSince = (Get-Date).AddDays(-7)
$failedLogons = @(Get-WinEvent -FilterHashtable @{LogName='Security'; Id=4625; StartTime=$failedSince} -MaxEvents 100 -ErrorAction SilentlyContinue)
if ($failedLogons.Count -gt 0) {
    # Group by target account
    $grouped = $failedLogons | ForEach-Object {
        $xml = [xml]$_.ToXml()
        $targetUser = ($xml.Event.EventData.Data | Where-Object { $_.Name -eq 'TargetUserName' }).'#text'
        $sourceIP = ($xml.Event.EventData.Data | Where-Object { $_.Name -eq 'IpAddress' }).'#text'
        [PSCustomObject]@{ Time = $_.TimeCreated; User = $targetUser; SourceIP = $sourceIP }
    }
    $summary = $grouped | Group-Object User | Sort-Object Count -Descending | Select-Object Count, Name
    Write-Host "  $($failedLogons.Count) failed logon attempts in last 7 days:" -ForegroundColor Yellow
    $summary | ForEach-Object { Write-Host "    $($_.Name): $($_.Count) failures" -ForegroundColor Yellow }

    $audit.FailedLogons = [ordered]@{
        TotalCount = $failedLogons.Count
        ByUser = @($summary | ForEach-Object { [ordered]@{ User = $_.Name; Count = $_.Count } })
        Recent = @($grouped | Select-Object -First 20 | ForEach-Object {
            [ordered]@{ Time = $_.Time.ToString("yyyy-MM-dd HH:mm:ss"); User = $_.User; SourceIP = $_.SourceIP }
        })
    }
} else {
    Write-Host "  No failed logon attempts" -ForegroundColor Green
    $audit.FailedLogons = [ordered]@{ TotalCount = 0 }
}
Write-Host "  [OK] Failed logons collected" -ForegroundColor Green

# =====================================================
# 45. SECURITY QUICK CHECKS
# =====================================================
Write-Host ""
Write-Host "=== 45. SECURITY QUICK CHECKS ===" -ForegroundColor Cyan
$audit.SecurityChecks = [ordered]@{}

# --- Password Policy ---
Write-Host "  Running: Get-ADDefaultDomainPasswordPolicy"
try {
    $pp = Get-ADDefaultDomainPasswordPolicy -ErrorAction Stop
    $passPolicy = [ordered]@{
        MinLength            = $pp.MinPasswordLength
        ComplexityEnabled    = $pp.ComplexityEnabled
        MaxAge               = "$($pp.MaxPasswordAge)"
        MinAge               = "$($pp.MinPasswordAge)"
        HistoryCount         = $pp.PasswordHistoryCount
        LockoutThreshold     = $pp.LockoutThreshold
        LockoutDuration      = "$($pp.LockoutDuration)"
        LockoutWindow        = "$($pp.LockoutObservationWindow)"
        ReversibleEncryption = $pp.ReversibleEncryptionEnabled
    }
    $passPolicy.GetEnumerator() | ForEach-Object { Write-Host "    $($_.Key): $($_.Value)" }

    # Flag weak settings
    $passIssues = @()
    if ($pp.MinPasswordLength -lt 12) { $passIssues += "Min password length is $($pp.MinPasswordLength) - recommend 12+" }
    if (-not $pp.ComplexityEnabled) { $passIssues += "Complexity is DISABLED" }
    if ($pp.LockoutThreshold -eq 0) { $passIssues += "Account lockout is DISABLED - brute force risk" }
    if ($pp.ReversibleEncryptionEnabled) { $passIssues += "Reversible encryption is ENABLED - critical risk" }
    if ($pp.PasswordHistoryCount -lt 12) { $passIssues += "Password history is only $($pp.PasswordHistoryCount) - recommend 12+" }

    $passPolicy.Issues = $passIssues
    if ($passIssues) {
        $passIssues | ForEach-Object { Write-Host "    [WARN] $_" -ForegroundColor Yellow }
    }
    $audit.SecurityChecks.PasswordPolicy = $passPolicy
    Write-Host "  [OK] Password policy checked" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "PasswordPolicy"; Error = $_.Exception.Message }
}

# --- Fine-Grained Password Policies ---
Write-Host "  Running: Get-ADFineGrainedPasswordPolicy"
try {
    $fgpp = Get-ADFineGrainedPasswordPolicy -Filter * -ErrorAction Stop
    if ($fgpp) {
        $audit.SecurityChecks.FineGrainedPolicies = @($fgpp | ForEach-Object {
            Write-Host "    FGPP: $($_.Name) - MinLen=$($_.MinPasswordLength), MaxAge=$($_.MaxPasswordAge), AppliesTo=$($_.AppliesTo -join ', ')"
            [ordered]@{
                Name = $_.Name; Precedence = $_.Precedence; MinLength = $_.MinPasswordLength
                MaxAge = "$($_.MaxPasswordAge)"; LockoutThreshold = $_.LockoutThreshold
                AppliesTo = @($_.AppliesTo)
            }
        })
        Write-Host "  [OK] Fine-grained policies checked" -ForegroundColor Green
    } else {
        Write-Host "    No fine-grained password policies configured"
        $audit.SecurityChecks.FineGrainedPolicies = @()
    }
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "FineGrainedPolicies"; Error = $_.Exception.Message }
}

# --- LAPS Status ---
Write-Host "  Running: LAPS check"
try {
    $lapsCheck = [ordered]@{ LegacyLAPS = $false; WindowsLAPS = $false }

    # Check for Legacy LAPS schema extension
    $schemaNC = (Get-ADRootDSE).schemaNamingContext
    $lapsSchema = Get-ADObject -SearchBase $schemaNC -Filter "Name -eq 'ms-Mcs-AdmPwd'" -ErrorAction SilentlyContinue
    if ($lapsSchema) {
        $lapsCheck.LegacyLAPS = $true
        Write-Host "    Legacy LAPS (ms-Mcs-AdmPwd): INSTALLED" -ForegroundColor Green
        # Count computers with LAPS passwords
        $lapsComputers = (Get-ADComputer -Filter * -Properties ms-Mcs-AdmPwd | Where-Object { $_.'ms-Mcs-AdmPwd' }).Count
        $totalComputers = (Get-ADComputer -Filter *).Count
        $lapsCheck.LegacyLAPSCoverage = "$lapsComputers / $totalComputers computers"
        Write-Host "    Coverage: $lapsComputers / $totalComputers computers have LAPS passwords"
    } else {
        Write-Host "    Legacy LAPS: NOT INSTALLED" -ForegroundColor Yellow
    }

    # Check for Windows LAPS schema
    $winLapsSchema = Get-ADObject -SearchBase $schemaNC -Filter "Name -eq 'ms-LAPS-Password'" -ErrorAction SilentlyContinue
    if ($winLapsSchema) {
        $lapsCheck.WindowsLAPS = $true
        Write-Host "    Windows LAPS (ms-LAPS-Password): INSTALLED" -ForegroundColor Green
    } else {
        Write-Host "    Windows LAPS: NOT INSTALLED" -ForegroundColor Yellow
    }

    if (-not $lapsCheck.LegacyLAPS -and -not $lapsCheck.WindowsLAPS) {
        Write-Host "    [WARN] No LAPS deployment detected - local admin passwords are likely shared/static" -ForegroundColor Red
    }
    $audit.SecurityChecks.LAPS = $lapsCheck
    Write-Host "  [OK] LAPS checked" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "LAPS"; Error = $_.Exception.Message }
}

# --- SMBv1 + SMB Signing ---
Write-Host "  Running: SMB configuration check"
try {
    $smbConfig = Get-SmbServerConfiguration -ErrorAction Stop

    # SMBv1
    if ($smbConfig.EnableSMB1Protocol) {
        Write-Host "    [WARN] SMBv1 is ENABLED - known vulnerability, should be disabled" -ForegroundColor Red
    } else {
        Write-Host "    SMBv1: Disabled" -ForegroundColor Green
    }
    $audit.SecurityChecks.SMBv1Enabled = $smbConfig.EnableSMB1Protocol

    # SMB Signing
    $smbSigning = [ordered]@{
        RequireSecuritySignature = $smbConfig.RequireSecuritySignature
        EnableSecuritySignature  = $smbConfig.EnableSecuritySignature
        EncryptData              = $smbConfig.EncryptData
    }
    $smbSigning.GetEnumerator() | ForEach-Object { Write-Host "    $($_.Key): $($_.Value)" }
    if (-not $smbConfig.RequireSecuritySignature) {
        Write-Host "    [WARN] SMB signing is NOT required - relay attack risk" -ForegroundColor Yellow
    }
    $audit.SecurityChecks.SMBSigning = $smbSigning
    Write-Host "  [OK] SMB config checked" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "SMBConfig"; Error = $_.Exception.Message }
}

# --- Machine Account Quota ---
Write-Host "  Running: Machine Account Quota check"
try {
    $maq = (Get-ADObject -Identity $script:domain.DistinguishedName -Properties ms-DS-MachineAccountQuota).'ms-DS-MachineAccountQuota'
    Write-Host "    Machine Account Quota: $maq"
    if ($maq -gt 0) {
        Write-Host "    [WARN] Any domain user can join up to $maq computers to the domain" -ForegroundColor Yellow
    }
    $audit.SecurityChecks.MachineAccountQuota = $maq
    Write-Host "  [OK] Machine account quota checked" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "MachineAccountQuota"; Error = $_.Exception.Message }
}

# --- AD Recycle Bin ---
Write-Host "  Running: AD Recycle Bin check"
try {
    $recycleBin = Get-ADOptionalFeature -Filter "Name -eq 'Recycle Bin Feature'" -ErrorAction Stop
    if ($recycleBin.EnabledScopes.Count -gt 0) {
        Write-Host "    AD Recycle Bin: ENABLED" -ForegroundColor Green
        $audit.SecurityChecks.RecycleBinEnabled = $true
    } else {
        Write-Host "    [WARN] AD Recycle Bin is NOT enabled - accidental deletions are unrecoverable" -ForegroundColor Yellow
        $audit.SecurityChecks.RecycleBinEnabled = $false
    }
    Write-Host "  [OK] Recycle Bin checked" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "RecycleBin"; Error = $_.Exception.Message }
}

# --- Accounts with Password Never Expires ---
Write-Host "  Running: Accounts with PasswordNeverExpires"
try {
    $neverExpire = Get-ADUser -Filter { PasswordNeverExpires -eq $true -and Enabled -eq $true } -Properties PasswordNeverExpires, PasswordLastSet |
        Select-Object Name, SamAccountName, PasswordLastSet
    if ($neverExpire) {
        Write-Host "    [WARN] $($neverExpire.Count) enabled accounts have PasswordNeverExpires:" -ForegroundColor Yellow
        $neverExpire | ForEach-Object { Write-Host "      $($_.SamAccountName) - Password set: $($_.PasswordLastSet)" -ForegroundColor Yellow }
    } else {
        Write-Host "    No enabled accounts with PasswordNeverExpires" -ForegroundColor Green
    }
    $audit.SecurityChecks.PasswordNeverExpires = @($neverExpire | ForEach-Object {
        [ordered]@{ Name = $_.Name; SAM = $_.SamAccountName; PasswordLastSet = if ($_.PasswordLastSet) { $_.PasswordLastSet.ToString("yyyy-MM-dd") } else { $null } }
    })
    Write-Host "  [OK] PasswordNeverExpires checked" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "PasswordNeverExpires"; Error = $_.Exception.Message }
}

# --- Stale Passwords (not changed in 90+ days) ---
Write-Host "  Running: Stale passwords check - enabled accounts, 90+ days"
try {
    $staleDate = (Get-Date).AddDays(-90)
    $stalePasswords = Get-ADUser -Filter { Enabled -eq $true -and PasswordLastSet -lt $staleDate } -Properties PasswordLastSet |
        Select-Object Name, SamAccountName, PasswordLastSet | Sort-Object PasswordLastSet
    if ($stalePasswords) {
        Write-Host "    [WARN] $($stalePasswords.Count) enabled accounts have passwords older than 90 days:" -ForegroundColor Yellow
        $stalePasswords | Select-Object -First 20 | ForEach-Object {
            Write-Host "      $($_.SamAccountName) - Last set: $($_.PasswordLastSet)" -ForegroundColor Yellow
        }
        if ($stalePasswords.Count -gt 20) { Write-Host "      ... and $($stalePasswords.Count - 20) more" -ForegroundColor Yellow }
    } else {
        Write-Host "    All enabled accounts have passwords changed within 90 days" -ForegroundColor Green
    }
    $audit.SecurityChecks.StalePasswords = @($stalePasswords | ForEach-Object {
        [ordered]@{ Name = $_.Name; SAM = $_.SamAccountName; PasswordLastSet = if ($_.PasswordLastSet) { $_.PasswordLastSet.ToString("yyyy-MM-dd") } else { $null } }
    })
    Write-Host "  [OK] Stale passwords checked" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "StalePasswords"; Error = $_.Exception.Message }
}

# --- Inactive Accounts (not logged in 90+ days) ---
Write-Host "  Running: Inactive accounts check - enabled, no logon 90+ days"
try {
    $inactiveDate = (Get-Date).AddDays(-90)
    $inactive = Get-ADUser -Filter { Enabled -eq $true -and LastLogonDate -lt $inactiveDate } -Properties LastLogonDate |
        Select-Object Name, SamAccountName, LastLogonDate | Sort-Object LastLogonDate
    if ($inactive) {
        Write-Host "    [WARN] $($inactive.Count) enabled accounts have not logged in for 90+ days:" -ForegroundColor Yellow
        $inactive | Select-Object -First 20 | ForEach-Object {
            Write-Host "      $($_.SamAccountName) - Last logon: $($_.LastLogonDate)" -ForegroundColor Yellow
        }
        if ($inactive.Count -gt 20) { Write-Host "      ... and $($inactive.Count - 20) more" -ForegroundColor Yellow }
    } else {
        Write-Host "    All enabled accounts have logged in within 90 days" -ForegroundColor Green
    }
    $audit.SecurityChecks.InactiveAccounts = @($inactive | ForEach-Object {
        [ordered]@{ Name = $_.Name; SAM = $_.SamAccountName; LastLogon = if ($_.LastLogonDate) { $_.LastLogonDate.ToString("yyyy-MM-dd") } else { $null } }
    })
    Write-Host "  [OK] Inactive accounts checked" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "InactiveAccounts"; Error = $_.Exception.Message }
}

# --- Locked Accounts ---
Write-Host "  Running: Locked accounts check"
try {
    $locked = Search-ADAccount -LockedOut | Where-Object { $_.Enabled } | Select-Object Name, SamAccountName, LastLogonDate
    if ($locked) {
        Write-Host "    [WARN] $($locked.Count) accounts are currently LOCKED OUT:" -ForegroundColor Yellow
        $locked | ForEach-Object { Write-Host "      $($_.SamAccountName)" -ForegroundColor Yellow }
    } else {
        Write-Host "    No locked accounts" -ForegroundColor Green
    }
    $audit.SecurityChecks.LockedAccounts = @($locked | ForEach-Object {
        [ordered]@{ Name = $_.Name; SAM = $_.SamAccountName }
    })
    Write-Host "  [OK] Locked accounts checked" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "LockedAccounts"; Error = $_.Exception.Message }
}

# --- AdminSDHolder / Protected Users ---
Write-Host "  Running: Protected Users group check"
try {
    $protectedMembers = Get-ADGroupMember "Protected Users" -ErrorAction Stop | Select-Object Name, SamAccountName
    if ($protectedMembers) {
        Write-Host "    Protected Users members:" -ForegroundColor Green
        $protectedMembers | ForEach-Object { Write-Host "      $($_.SamAccountName)" }
    } else {
        Write-Host "    [WARN] Protected Users group is EMPTY - admin accounts should be added" -ForegroundColor Yellow
    }
    $audit.SecurityChecks.ProtectedUsers = @($protectedMembers | ForEach-Object {
        [ordered]@{ Name = $_.Name; SAM = $_.SamAccountName }
    })
    Write-Host "  [OK] Protected Users checked" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "ProtectedUsers"; Error = $_.Exception.Message }
}

# --- AdminSDHolder protected accounts ---
Write-Host "  Running: AdminSDHolder check"
try {
    $adminSDHolder = Get-ADUser -Filter { AdminCount -eq 1 -and Enabled -eq $true } -Properties AdminCount, MemberOf |
        Select-Object Name, SamAccountName
    if ($adminSDHolder) {
        Write-Host "    Accounts with AdminCount=1 - $($adminSDHolder.Count) accounts:" -ForegroundColor Yellow
        $adminSDHolder | ForEach-Object { Write-Host "      $($_.SamAccountName)" }
    }
    $audit.SecurityChecks.AdminSDHolder = @($adminSDHolder | ForEach-Object {
        [ordered]@{ Name = $_.Name; SAM = $_.SamAccountName }
    })
    Write-Host "  [OK] AdminSDHolder checked" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "AdminSDHolder"; Error = $_.Exception.Message }
}

# --- Kerberoastable Accounts (SPNs on user accounts) ---
Write-Host "  Running: Kerberoastable accounts check"
try {
    $spnUsers = Get-ADUser -Filter { ServicePrincipalName -like "*" -and Enabled -eq $true } -Properties ServicePrincipalName, PasswordLastSet |
        Where-Object { $_.SamAccountName -ne 'krbtgt' } |
        Select-Object Name, SamAccountName, PasswordLastSet, ServicePrincipalName
    if ($spnUsers) {
        Write-Host "    [WARN] $($spnUsers.Count) enabled user accounts have SPNs - kerberoastable:" -ForegroundColor Yellow
        $spnUsers | ForEach-Object {
            Write-Host "      $($_.SamAccountName) - Password set: $($_.PasswordLastSet) - SPN: $($_.ServicePrincipalName[0])" -ForegroundColor Yellow
        }
    } else {
        Write-Host "    No kerberoastable user accounts found" -ForegroundColor Green
    }
    $audit.SecurityChecks.Kerberoastable = @($spnUsers | ForEach-Object {
        [ordered]@{
            Name = $_.Name; SAM = $_.SamAccountName
            PasswordLastSet = if ($_.PasswordLastSet) { $_.PasswordLastSet.ToString("yyyy-MM-dd") } else { $null }
            SPNs = @($_.ServicePrincipalName)
        }
    })
    Write-Host "  [OK] Kerberoastable checked" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "Kerberoastable"; Error = $_.Exception.Message }
}

# --- ASREPRoastable Accounts (PreAuth disabled) ---
Write-Host "  Running: ASREPRoastable accounts check"
try {
    $asrep = Get-ADUser -Filter { DoesNotRequirePreAuth -eq $true -and Enabled -eq $true } -Properties DoesNotRequirePreAuth |
        Select-Object Name, SamAccountName
    if ($asrep) {
        Write-Host "    [WARN] $($asrep.Count) accounts do NOT require Kerberos pre-auth - ASREPRoastable:" -ForegroundColor Red
        $asrep | ForEach-Object { Write-Host "      $($_.SamAccountName)" -ForegroundColor Red }
    } else {
        Write-Host "    No ASREPRoastable accounts found" -ForegroundColor Green
    }
    $audit.SecurityChecks.ASREPRoastable = @($asrep | ForEach-Object {
        [ordered]@{ Name = $_.Name; SAM = $_.SamAccountName }
    })
    Write-Host "  [OK] ASREPRoastable checked" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "ASREPRoastable"; Error = $_.Exception.Message }
}

# --- NULL Sessions ---
Write-Host "  Running: NULL session check"
try {
    $lsaReg = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' -ErrorAction Stop
    $restrictAnon = $lsaReg.RestrictAnonymous
    $restrictAnonSAM = $lsaReg.RestrictAnonymousSAM
    $everyoneAnon = $lsaReg.EveryoneIncludesAnonymous

    $nullSession = [ordered]@{
        RestrictAnonymous           = $restrictAnon
        RestrictAnonymousSAM        = $restrictAnonSAM
        EveryoneIncludesAnonymous   = $everyoneAnon
    }
    $nullSession.GetEnumerator() | ForEach-Object { Write-Host "    $($_.Key): $($_.Value)" }

    if ($everyoneAnon -eq 1) {
        Write-Host "    [WARN] EveryoneIncludesAnonymous is ENABLED - anonymous users have Everyone access" -ForegroundColor Red
    }
    if ($restrictAnon -eq 0) {
        Write-Host "    [WARN] RestrictAnonymous is 0 - anonymous enumeration of shares/accounts possible" -ForegroundColor Yellow
    }

    $audit.SecurityChecks.NullSessions = $nullSession
    Write-Host "  [OK] NULL sessions checked" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "NullSessions"; Error = $_.Exception.Message }
}

# --- Insecure DNS Zones ---
Write-Host "  Running: Insecure DNS zones check"
try {
    Import-Module DnsServer -ErrorAction Stop
    $insecureZones = Get-DnsServerZone | Where-Object { $_.DynamicUpdate -eq 'NonsecureAndSecure' -and -not $_.IsAutoCreated }
    if ($insecureZones) {
        Write-Host "    [WARN] DNS zones allowing nonsecure dynamic updates:" -ForegroundColor Yellow
        $insecureZones | ForEach-Object { Write-Host "      $($_.ZoneName)" -ForegroundColor Yellow }
    } else {
        Write-Host "    All DNS zones use secure-only dynamic updates" -ForegroundColor Green
    }
    $audit.SecurityChecks.InsecureDNSZones = @($insecureZones | ForEach-Object { $_.ZoneName })
    Write-Host "  [OK] DNS zone security checked" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "InsecureDNSZones"; Error = $_.Exception.Message }
}

# --- SYSVOL Password Search (GPP cpassword) ---
Write-Host "  Running: SYSVOL GPP password search"
try {
    $dnsDomain = if ($env:USERDNSDOMAIN) { $env:USERDNSDOMAIN } elseif ($script:domain) { $script:domain.DNSRoot } else { (Get-ADDomain).DNSRoot }
    $sysvolPath = "\\$dnsDomain\SYSVOL\$dnsDomain"
    $gppFiles = Get-ChildItem -Path $sysvolPath -Recurse -Include "*.xml" -ErrorAction Stop
    $cpasswordFiles = @()
    foreach ($file in $gppFiles) {
        $content = Get-Content $file.FullName -ErrorAction SilentlyContinue
        if ($content -match 'cpassword') {
            $cpasswordFiles += $file.FullName
            Write-Host "    [CRITICAL] GPP password found in: $($file.FullName)" -ForegroundColor Red
        }
    }
    if ($cpasswordFiles.Count -eq 0) {
        Write-Host "    No GPP passwords found in SYSVOL" -ForegroundColor Green
    }
    $audit.SecurityChecks.GPPPasswords = $cpasswordFiles
    Write-Host "  [OK] SYSVOL password search complete" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "GPPPasswords"; Error = $_.Exception.Message }
}

# --- Old/EOL Operating Systems ---
Write-Host "  Running: End-of-life OS check"
try {
    $eolPatterns = @('Windows XP*', 'Windows Vista*', 'Windows 7*', 'Windows 8 *', 'Windows 8.1*',
        'Windows Server 2003*', 'Windows Server 2008*', 'Windows Server 2012 *', 'Windows Server 2012 R2*')
    $eolComputers = @()
    foreach ($pattern in $eolPatterns) {
        $found = Get-ADComputer -Filter "OperatingSystem -like '$pattern' -and Enabled -eq 'True'" -Properties OperatingSystem, LastLogonDate -ErrorAction SilentlyContinue
        if ($found) { $eolComputers += $found }
    }
    if ($eolComputers) {
        Write-Host "    [WARN] $($eolComputers.Count) enabled computers running END OF LIFE operating systems:" -ForegroundColor Red
        $eolComputers | ForEach-Object { Write-Host "      $($_.Name) - $($_.OperatingSystem) - Last logon: $($_.LastLogonDate)" -ForegroundColor Red }
    } else {
        Write-Host "    No end-of-life operating systems detected" -ForegroundColor Green
    }
    $audit.SecurityChecks.EOLComputers = @($eolComputers | ForEach-Object {
        [ordered]@{ Name = $_.Name; OS = $_.OperatingSystem; LastLogon = if ($_.LastLogonDate) { $_.LastLogonDate.ToString("yyyy-MM-dd") } else { $null } }
    })
    Write-Host "  [OK] EOL OS check complete" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "EOLComputers"; Error = $_.Exception.Message }
}

# --- Functional Level ---
Write-Host "  Running: Functional level assessment"
try {
    $domainFL = if ($script:domain) { $script:domain.DomainMode } else { (Get-ADDomain).DomainMode }
    $forestFL = if ($script:forest) { $script:forest.ForestMode } else { (Get-ADForest).ForestMode }
    Write-Host "    Domain Functional Level: $domainFL"
    Write-Host "    Forest Functional Level: $forestFL"

    $outdatedLevels = @('Windows2003Domain', 'Windows2008Domain', 'Windows2008R2Domain', 'Windows2012Domain', 'Windows2012R2Domain',
        'Windows2003Forest', 'Windows2008Forest', 'Windows2008R2Forest', 'Windows2012Forest', 'Windows2012R2Forest')
    if ($domainFL -in $outdatedLevels -or $forestFL -in $outdatedLevels) {
        Write-Host "    [WARN] Functional level is outdated - limits security features like Protected Users, auth silos" -ForegroundColor Yellow
    }

    $audit.SecurityChecks.FunctionalLevel = [ordered]@{ Domain = "$domainFL"; Forest = "$forestFL" }
    Write-Host "  [OK] Functional level checked" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "FunctionalLevel"; Error = $_.Exception.Message }
}

# --- Replication Type (FRS vs DFSR) ---
Write-Host "  Running: SYSVOL replication type check"
try {
    $dn = if ($script:domain) { $script:domain.DistinguishedName } else { (Get-ADDomain).DistinguishedName }
    $replType = (Get-ADObject "CN=DFSR-GlobalSettings,CN=System,$dn" -ErrorAction Stop)
    Write-Host "    SYSVOL Replication: DFSR" -ForegroundColor Green
    $audit.SecurityChecks.SYSVOLReplication = "DFSR"
}
catch {
    Write-Host "    SYSVOL Replication: FRS (legacy - should migrate to DFSR)" -ForegroundColor Yellow
    $audit.SecurityChecks.SYSVOLReplication = "FRS"
}
Write-Host "  [OK] Replication type checked" -ForegroundColor Green

# --- LDAP Signing + Channel Binding ---
Write-Host "  Running: LDAP signing and channel binding check"
try {
    $ntdsReg = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Parameters' -ErrorAction Stop

    $ldapSigning = $ntdsReg.LDAPServerIntegrity
    $ldapSigningText = switch ($ldapSigning) {
        0 { "None (not required)" } 1 { "Require signing" } 2 { "Require signing" } default { "Unknown ($ldapSigning)" }
    }
    Write-Host "    LDAP Server Signing: $ldapSigningText"
    if ($ldapSigning -eq 0 -or $null -eq $ldapSigning) {
        Write-Host "    [WARN] LDAP signing is NOT required - MITM risk" -ForegroundColor Yellow
    }
    $audit.SecurityChecks.LDAPSigning = $ldapSigningText

    $ldapCB = $ntdsReg.LdapEnforceChannelBinding
    $ldapCBText = switch ($ldapCB) {
        0 { "Never" } 1 { "When supported" } 2 { "Always" } default { "Not configured (0)" }
    }
    Write-Host "    LDAP Channel Binding: $ldapCBText"
    if ($ldapCB -ne 2) {
        Write-Host "    [WARN] LDAP channel binding is not set to Always - credential relay risk" -ForegroundColor Yellow
    }
    $audit.SecurityChecks.LDAPChannelBinding = $ldapCBText
    Write-Host "  [OK] LDAP settings checked" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "LDAPSettings"; Error = $_.Exception.Message }
}

# --- DCs Not Owned by Domain Admins ---
Write-Host "  Running: DC ownership check"
try {
    $dcsNotOwned = Get-ADComputer -Filter { PrimaryGroupID -eq 516 } -Properties nTSecurityDescriptor -ErrorAction Stop | ForEach-Object {
        $owner = $_.nTSecurityDescriptor.Owner
        if ($owner -notmatch 'Domain Admins') {
            [PSCustomObject]@{ Name = $_.Name; Owner = $owner }
        }
    }
    if ($dcsNotOwned) {
        Write-Host "    [WARN] DCs not owned by Domain Admins:" -ForegroundColor Yellow
        $dcsNotOwned | ForEach-Object { Write-Host "      $($_.Name) owned by $($_.Owner)" -ForegroundColor Yellow }
    } else {
        Write-Host "    All DCs owned by Domain Admins" -ForegroundColor Green
    }
    $audit.SecurityChecks.DCsNotOwnedByDA = @($dcsNotOwned | ForEach-Object {
        [ordered]@{ Name = $_.Name; Owner = $_.Owner }
    })
    Write-Host "  [OK] DC ownership checked" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "DCOwnership"; Error = $_.Exception.Message }
}

# --- Recent Changes (last 30 days) ---
Write-Host "  Running: Recent AD changes - last 30 days"
try {
    $recentDate = (Get-Date).AddDays(-30)
    $newUsers = Get-ADUser -Filter { WhenCreated -gt $recentDate } -Properties WhenCreated |
        Select-Object Name, SamAccountName, WhenCreated
    $newGroups = Get-ADGroup -Filter { WhenCreated -gt $recentDate } -Properties WhenCreated |
        Select-Object Name, WhenCreated
    $newComputers = Get-ADComputer -Filter { WhenCreated -gt $recentDate } -Properties WhenCreated |
        Select-Object Name, WhenCreated

    if ($newUsers) {
        Write-Host "    New users in last 30 days:" -ForegroundColor Yellow
        $newUsers | ForEach-Object { Write-Host "      $($_.SamAccountName) - Created: $($_.WhenCreated)" }
    }
    if ($newGroups) {
        Write-Host "    New groups in last 30 days:" -ForegroundColor Yellow
        $newGroups | ForEach-Object { Write-Host "      $($_.Name) - Created: $($_.WhenCreated)" }
    }
    if ($newComputers) {
        Write-Host "    New computers in last 30 days:" -ForegroundColor Yellow
        $newComputers | ForEach-Object { Write-Host "      $($_.Name) - Created: $($_.WhenCreated)" }
    }
    if (-not $newUsers -and -not $newGroups -and -not $newComputers) {
        Write-Host "    No new AD objects in the last 30 days" -ForegroundColor Green
    }

    $audit.SecurityChecks.RecentChanges = [ordered]@{
        NewUsers = @($newUsers | ForEach-Object { [ordered]@{ Name = $_.Name; SAM = $_.SamAccountName; Created = $_.WhenCreated.ToString("yyyy-MM-dd") } })
        NewGroups = @($newGroups | ForEach-Object { [ordered]@{ Name = $_.Name; Created = $_.WhenCreated.ToString("yyyy-MM-dd") } })
        NewComputers = @($newComputers | ForEach-Object { [ordered]@{ Name = $_.Name; Created = $_.WhenCreated.ToString("yyyy-MM-dd") } })
    }
    Write-Host "  [OK] Recent changes checked" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "RecentChanges"; Error = $_.Exception.Message }
}

# --- krbtgt Password Age ---
Write-Host "  Running: krbtgt password age check"
try {
    $krbtgt = Get-ADUser 'krbtgt' -Properties PasswordLastSet -ErrorAction Stop
    $krbtgtAge = ((Get-Date) - $krbtgt.PasswordLastSet).Days
    Write-Host "    krbtgt password last set: $($krbtgt.PasswordLastSet) - $krbtgtAge days ago"
    if ($krbtgtAge -gt 180) {
        Write-Host "    [WARN] krbtgt password is $krbtgtAge days old - should be rotated at least every 180 days" -ForegroundColor Yellow
    }
    $audit.SecurityChecks.KrbtgtPasswordAge = [ordered]@{
        LastSet = $krbtgt.PasswordLastSet.ToString("yyyy-MM-dd")
        AgeDays = $krbtgtAge
    }
    Write-Host "  [OK] krbtgt checked" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "KrbtgtAge"; Error = $_.Exception.Message }
}

# --- Audit Policy (are logons being audited?) ---
Write-Host "  Running: Audit policy check"
try {
    $auditPolicy = auditpol /get /category:* 2>&1
    $auditPolicy | ForEach-Object { Write-Host "    $_" }

    # Check critical audit settings
    $audit.SecurityChecks.AuditPolicyRaw = @($auditPolicy | ForEach-Object { "$_".Trim() } | Where-Object { $_ })
    Write-Host "  [OK] Audit policy collected" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "AuditPolicy"; Error = $_.Exception.Message }
}

# --- Stopped Auto-Start Services ---
Write-Host "  Running: Auto-start services that are stopped"
try {
    $stoppedAuto = Get-Service | Where-Object { $_.StartType -eq 'Automatic' -and $_.Status -ne 'Running' } |
        Select-Object Name, DisplayName, Status, StartType
    if ($stoppedAuto) {
        Write-Host "    [WARN] $($stoppedAuto.Count) auto-start services are NOT running:" -ForegroundColor Yellow
        $stoppedAuto | Format-Table -AutoSize
    } else {
        Write-Host "    All auto-start services are running" -ForegroundColor Green
    }
    $audit.SecurityChecks.StoppedAutoServices = @($stoppedAuto | ForEach-Object {
        [ordered]@{ Name = $_.Name; DisplayName = $_.DisplayName; Status = "$($_.Status)" }
    })
    Write-Host "  [OK] Service state check complete" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $audit._errors += @{ Section = "StoppedAutoServices"; Error = $_.Exception.Message }
}

Write-Host ""
Write-Host "  === Security Summary ===" -ForegroundColor Cyan
$totalWarnings = 0
if ($audit.SecurityChecks.PasswordPolicy.Issues) { $totalWarnings += $audit.SecurityChecks.PasswordPolicy.Issues.Count }
if ($audit.SecurityChecks.SMBv1Enabled) { $totalWarnings++ }
if ($audit.SecurityChecks.MachineAccountQuota -gt 0) { $totalWarnings++ }
if (-not $audit.SecurityChecks.RecycleBinEnabled) { $totalWarnings++ }
if ($audit.SecurityChecks.PasswordNeverExpires.Count -gt 0) { $totalWarnings++ }
if ($audit.SecurityChecks.Kerberoastable.Count -gt 0) { $totalWarnings++ }
if ($audit.SecurityChecks.ASREPRoastable.Count -gt 0) { $totalWarnings++ }
if ($audit.SecurityChecks.GPPPasswords.Count -gt 0) { $totalWarnings++ }
if ($audit.SecurityChecks.EOLComputers.Count -gt 0) { $totalWarnings++ }
if (-not $audit.SecurityChecks.LAPS.LegacyLAPS -and -not $audit.SecurityChecks.LAPS.WindowsLAPS) { $totalWarnings++ }

if ($totalWarnings -gt 0) {
    Write-Host "  $totalWarnings security findings detected - review warnings above" -ForegroundColor Yellow
} else {
    Write-Host "  No critical security findings" -ForegroundColor Green
}

# =====================================================
# DONE - SAVE JSON
# =====================================================
Write-Host ""
Write-Host "======================================="
Write-Host "      AUDIT COMPLETE"
Write-Host "======================================="

$errorCount = $audit._errors.Count
if ($errorCount -gt 0) {
    Write-Host "Completed with $errorCount section errors (see _errors in JSON)" -ForegroundColor Yellow
} else {
    Write-Host "All sections completed successfully" -ForegroundColor Green
}

Write-Host "Transcript: $TxtFile"
Write-Host "JSON data:  $JsonFile"
Write-Host "======================================="

Stop-Transcript

# Save JSON
$audit | ConvertTo-Json -Depth 10 | Out-File $JsonFile -Encoding UTF8
