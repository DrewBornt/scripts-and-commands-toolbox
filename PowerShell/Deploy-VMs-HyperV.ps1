#Requires -Modules Hyper-V
#Requires -RunAsAdministrator

# ============================================================
# CLIENT CONFIGURATION — EDIT THESE VALUES BEFORE RUNNING
# ============================================================

# Full path to your sysprepped golden image VHDX
$GoldenImagePath    = "D:\GoldenImages\WS2022-Gold.vhdx"         # e.g. "D:\GoldenImages\WS2022-Gold.vhdx"

# Root folder where VM subdirectories will be created
# Script will create: $VMStorageRoot\<VMName>\<VMName>.vhdx
$VMStorageRoot      = "D:\VMs"                                    # e.g. "D:\VMs"

# Hyper-V virtual switch name — must match exactly as shown in Hyper-V Manager
$vSwitchName        = "Production"                                # e.g. "Production" or "External Switch"

# Default VM hardware — operator can override these at runtime
$DefaultRAMBytes    = 4GB                                         # e.g. 4GB, 8GB, 16GB
$DefaultCPUCount    = 4                                           # e.g. 2, 4, 8

# Active Directory domain name
$DomainName         = "DOMAIN.LOCAL"                              # e.g. "contoso.local"

# OU distinguished name — computer object will be created here at domain join
$DomainOU           = "OU=Servers,OU=Company,DC=DOMAIN,DC=LOCAL"  # e.g. "OU=Servers,OU=Contoso,DC=contoso,DC=local"

# Domain join service account username (password prompted securely at runtime)
$JoinAccount        = "svc-domainjoin"                            # e.g. "svc-domainjoin"

# DNS servers
$PrimaryDNS         = "10.0.0.1"                                  # e.g. "192.168.1.10"
$SecondaryDNS       = "10.0.0.2"                                  # e.g. "192.168.1.11" — leave "" if none

# Default subnet prefix length (e.g. 24 = /24 = 255.255.255.0)
$DefaultPrefixLength = 24                                         # e.g. 24, 16, 8

# Default gateway — can be overridden at runtime
$DefaultGateway     = "10.0.0.254"                                # e.g. "192.168.1.1"

# ============================================================
# END OF CLIENT CONFIGURATION — DO NOT EDIT BELOW THIS LINE
# ============================================================


# ============================================================
# FUNCTIONS
# ============================================================

function Write-Header {
    Clear-Host
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "  Hyper-V VM Deployment Script" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Prompt-WithDefault {
    param(
        [string]$Message,
        [string]$Default
    )
    $input = Read-Host "$Message [default: $Default]"
    if ([string]::IsNullOrWhiteSpace($input)) {
        return $Default
    }
    return $input
}

function Validate-IPAddress {
    param([string]$IP)
    $parsed = $null
    return [System.Net.IPAddress]::TryParse($IP, [ref]$parsed)
}

function Inject-UnattendXml {
    param(
        [string]$VHDXPath,
        [string]$VMName,
        [string]$IPAddress,
        [int]$PrefixLength,
        [string]$Gateway,
        [string]$JoinPassword
    )

    Write-Host "`n[*] Generating unattend.xml..." -ForegroundColor Yellow

    # Build DNS entries
    $dnsEntries = "                    <IpAddress wcm:action=`"add`" wcm:keyValue=`"1`">$PrimaryDNS</IpAddress>"
    if (-not [string]::IsNullOrWhiteSpace($SecondaryDNS)) {
        $dnsEntries += "`n                    <IpAddress wcm:action=`"add`" wcm:keyValue=`"2`">$SecondaryDNS</IpAddress>"
    }

    $unattendXml = @"
<?xml version="1.0" encoding="utf-8"?>
<unattend xmlns="urn:schemas-microsoft-com:unattend">
  <settings pass="specialize">

    <!-- Computer Name -->
    <component name="Microsoft-Windows-Shell-Setup"
               processorArchitecture="amd64"
               publicKeyToken="31bf3856ad364e35"
               language="neutral"
               versionScope="nonSxS"
               xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State">
      <ComputerName>$VMName</ComputerName>
    </component>

    <!-- Static IP Configuration -->
    <component name="Microsoft-Windows-TCPIP"
               processorArchitecture="amd64"
               publicKeyToken="31bf3856ad364e35"
               language="neutral"
               versionScope="nonSxS"
               xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State">
      <Interfaces>
        <Interface wcm:action="add">
          <Identifier>Ethernet</Identifier>
          <Ipv4Settings>
            <DhcpEnabled>false</DhcpEnabled>
          </Ipv4Settings>
          <UnicastIpAddresses>
            <IpAddress wcm:action="add" wcm:keyValue="1">$IPAddress/$PrefixLength</IpAddress>
          </UnicastIpAddresses>
          <Routes>
            <Route wcm:action="add">
              <Identifier>1</Identifier>
              <NextHopAddress>$Gateway</NextHopAddress>
              <Prefix>0.0.0.0/0</Prefix>
            </Route>
          </Routes>
        </Interface>
      </Interfaces>
    </component>

    <!-- DNS Configuration -->
    <component name="Microsoft-Windows-DNS-Client"
               processorArchitecture="amd64"
               publicKeyToken="31bf3856ad364e35"
               language="neutral"
               versionScope="nonSxS"
               xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State">
      <Interfaces>
        <Interface wcm:action="add">
          <Identifier>Ethernet</Identifier>
          <DNSServerSearchOrder>
$dnsEntries
          </DNSServerSearchOrder>
          <DNSDomain>$DomainName</DNSDomain>
        </Interface>
      </Interfaces>
    </component>

    <!-- Domain Join -->
    <component name="Microsoft-Windows-UnattendedJoin"
               processorArchitecture="amd64"
               publicKeyToken="31bf3856ad364e35"
               language="neutral"
               versionScope="nonSxS"
               xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State">
      <Identification>
        <Credentials>
          <Domain>$DomainName</Domain>
          <Username>$JoinAccount</Username>
          <Password>$JoinPassword</Password>
        </Credentials>
        <JoinDomain>$DomainName</JoinDomain>
        <MachineObjectOU>$DomainOU</MachineObjectOU>
      </Identification>
    </component>

  </settings>

  <settings pass="oobeSystem">
    <component name="Microsoft-Windows-Shell-Setup"
               processorArchitecture="amd64"
               publicKeyToken="31bf3856ad364e35"
               language="neutral"
               versionScope="nonSxS"
               xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State">
      <OOBE>
        <HideEULAPage>true</HideEULAPage>
        <HideLocalAccountScreen>true</HideLocalAccountScreen>
        <HideOEMRegistrationScreen>true</HideOEMRegistrationScreen>
        <HideOnlineAccountScreens>true</HideOnlineAccountScreens>
        <HideWirelessSetupInOOBE>true</HideWirelessSetupInOOBE>
        <SkipMachineOOBE>true</SkipMachineOOBE>
        <SkipUserOOBE>true</SkipUserOOBE>
      </OOBE>
    </component>
  </settings>

</unattend>
"@

    # Write unattend.xml to a temp file
    $tempUnattend = "$env:TEMP\unattend_$VMName.xml"
    $unattendXml | Out-File -FilePath $tempUnattend -Encoding utf8 -Force

    Write-Host "[*] Mounting VHDX to inject unattend.xml..." -ForegroundColor Yellow

    try {
        # Mount the VHDX
        $mountResult = Mount-VHD -Path $VHDXPath -PassThru
        Start-Sleep -Seconds 3

        # Find the mounted drive letter (Windows partition — largest partition)
        $diskNumber = $mountResult.DiskNumber
        $partition = Get-Partition -DiskNumber $diskNumber | Where-Object { $_.Type -eq "Basic" } | Sort-Object Size -Descending | Select-Object -First 1

        if ($null -eq $partition) {
            throw "Could not find a Basic partition on the mounted VHDX."
        }

        # Assign a temporary drive letter if one isn't already assigned
        if ([string]::IsNullOrWhiteSpace($partition.DriveLetter) -or $partition.DriveLetter -eq "`0") {
            Add-PartitionAccessPath -DiskNumber $diskNumber -PartitionNumber $partition.PartitionNumber -AssignDriveLetter
            Start-Sleep -Seconds 2
            $partition = Get-Partition -DiskNumber $diskNumber -PartitionNumber $partition.PartitionNumber
        }

        $driveLetter = $partition.DriveLetter
        $panther = "${driveLetter}:\Windows\Panther"

        if (-not (Test-Path $panther)) {
            New-Item -ItemType Directory -Path $panther -Force | Out-Null
        }

        Copy-Item -Path $tempUnattend -Destination "$panther\unattend.xml" -Force
        Write-Host "[+] unattend.xml injected successfully into $panther" -ForegroundColor Green
    }
    catch {
        throw "Failed to inject unattend.xml: $_"
    }
    finally {
        # Always dismount, always clean up temp file
        Dismount-VHD -Path $VHDXPath -ErrorAction SilentlyContinue
        Remove-Item -Path $tempUnattend -Force -ErrorAction SilentlyContinue
        Write-Host "[*] VHDX dismounted." -ForegroundColor Yellow
    }
}


# ============================================================
# MAIN
# ============================================================

Write-Header

# --- Preflight checks ---
Write-Host "[*] Running preflight checks..." -ForegroundColor Yellow

if (-not (Test-Path $GoldenImagePath)) {
    Write-Host "[!] ERROR: Golden image not found at: $GoldenImagePath" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $VMStorageRoot)) {
    Write-Host "[!] ERROR: VM storage root not found at: $VMStorageRoot" -ForegroundColor Red
    exit 1
}

$switchExists = Get-VMSwitch -Name $vSwitchName -ErrorAction SilentlyContinue
if ($null -eq $switchExists) {
    Write-Host "[!] ERROR: vSwitch '$vSwitchName' not found. Check Hyper-V Manager." -ForegroundColor Red
    exit 1
}

Write-Host "[+] Preflight checks passed.`n" -ForegroundColor Green

# --- Gather inputs ---
Write-Host "--- VM Identity ---" -ForegroundColor Cyan
$VMName = ""
while ([string]::IsNullOrWhiteSpace($VMName)) {
    $VMName = Read-Host "VM Name (will also be used as hostname)"
    if ([string]::IsNullOrWhiteSpace($VMName)) {
        Write-Host "[!] VM name cannot be blank." -ForegroundColor Red
    }
}

# Check if VM already exists
if (Get-VM -Name $VMName -ErrorAction SilentlyContinue) {
    Write-Host "[!] ERROR: A VM named '$VMName' already exists on this host." -ForegroundColor Red
    exit 1
}

# Check if folder already exists
$VMFolder = Join-Path $VMStorageRoot $VMName
if (Test-Path $VMFolder) {
    Write-Host "[!] ERROR: Folder already exists at: $VMFolder" -ForegroundColor Red
    exit 1
}

Write-Host "`n--- Network Configuration ---" -ForegroundColor Cyan

$IPAddress = ""
while (-not (Validate-IPAddress $IPAddress)) {
    $IPAddress = Read-Host "Static IP Address"
    if (-not (Validate-IPAddress $IPAddress)) {
        Write-Host "[!] Invalid IP address. Please try again." -ForegroundColor Red
    }
}

$PrefixLength = Prompt-WithDefault -Message "Subnet Prefix Length" -Default $DefaultPrefixLength
$PrefixLength = [int]$PrefixLength

$Gateway = Prompt-WithDefault -Message "Default Gateway" -Default $DefaultGateway
while (-not (Validate-IPAddress $Gateway)) {
    Write-Host "[!] Invalid gateway address." -ForegroundColor Red
    $Gateway = Read-Host "Default Gateway"
}

Write-Host "`n--- VM Hardware ---" -ForegroundColor Cyan
$RAMInput = Prompt-WithDefault -Message "RAM in GB" -Default ($DefaultRAMBytes / 1GB)
$RAMBytes  = [int64]$RAMInput * 1GB
$CPUCount  = Prompt-WithDefault -Message "CPU Count" -Default $DefaultCPUCount
$CPUCount  = [int]$CPUCount

Write-Host "`n--- Domain Join Credentials ---" -ForegroundColor Cyan
Write-Host "Enter the password for domain join account: $JoinAccount@$DomainName" -ForegroundColor Gray
$JoinCred    = Get-Credential -UserName "$DomainName\$JoinAccount" -Message "Domain join credential for $DomainName"
$JoinPassword = $JoinCred.GetNetworkCredential().Password

# --- Summary before proceeding ---
Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  Deployment Summary — Review Before Continuing" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  VM Name       : $VMName"
Write-Host "  VM Folder     : $VMFolder"
Write-Host "  VHDX Source   : $GoldenImagePath"
Write-Host "  vSwitch       : $vSwitchName"
Write-Host "  RAM           : $($RAMBytes / 1GB) GB"
Write-Host "  CPUs          : $CPUCount"
Write-Host "  IP Address    : $IPAddress/$PrefixLength"
Write-Host "  Gateway       : $Gateway"
Write-Host "  Primary DNS   : $PrimaryDNS"
if (-not [string]::IsNullOrWhiteSpace($SecondaryDNS)) {
Write-Host "  Secondary DNS : $SecondaryDNS"
}
Write-Host "  Domain        : $DomainName"
Write-Host "  OU            : $DomainOU"
Write-Host "  Join Account  : $JoinAccount"
Write-Host "============================================================" -ForegroundColor Cyan

$confirm = Read-Host "`nProceed with deployment? (yes/no)"
if ($confirm -notmatch "^(yes|y)$") {
    Write-Host "`n[!] Deployment cancelled." -ForegroundColor Yellow
    exit 0
}

# --- Create VM folder ---
Write-Host "`n[*] Creating VM folder: $VMFolder" -ForegroundColor Yellow
New-Item -ItemType Directory -Path $VMFolder -Force | Out-Null
Write-Host "[+] Folder created." -ForegroundColor Green

# --- Copy golden image ---
$DestVHDX = Join-Path $VMFolder "$VMName.vhdx"
Write-Host "[*] Copying golden image — this may take a few minutes..." -ForegroundColor Yellow
Copy-Item -Path $GoldenImagePath -Destination $DestVHDX
Write-Host "[+] VHDX copied to: $DestVHDX" -ForegroundColor Green

# --- Inject unattend.xml ---
try {
    Inject-UnattendXml `
        -VHDXPath     $DestVHDX `
        -VMName       $VMName `
        -IPAddress    $IPAddress `
        -PrefixLength $PrefixLength `
        -Gateway      $Gateway `
        -JoinPassword $JoinPassword
}
catch {
    Write-Host "[!] ERROR during unattend.xml injection: $_" -ForegroundColor Red
    Write-Host "[!] Cleaning up..." -ForegroundColor Red
    Remove-Item -Path $VMFolder -Recurse -Force -ErrorAction SilentlyContinue
    exit 1
}

# Clear the join password from memory
$JoinPassword = $null

# --- Create the VM ---
Write-Host "`n[*] Creating VM in Hyper-V..." -ForegroundColor Yellow

try {
    $VM = New-VM `
        -Name              $VMName `
        -MemoryStartupBytes $RAMBytes `
        -VHDPath           $DestVHDX `
        -SwitchName        $vSwitchName `
        -Generation        2 `
        -Path              $VMStorageRoot

    Set-VM `
        -Name              $VMName `
        -ProcessorCount    $CPUCount `
        -DynamicMemory     `
        -MemoryMinimumBytes 512MB `
        -MemoryMaximumBytes ($RAMBytes * 2)

    # Generation 2 — ensure Secure Boot is on with Microsoft UEFI template
    Set-VMFirmware `
        -VMName            $VMName `
        -EnableSecureBoot  On `
        -SecureBootTemplate MicrosoftWindows

    Write-Host "[+] VM created successfully." -ForegroundColor Green
}
catch {
    Write-Host "[!] ERROR creating VM: $_" -ForegroundColor Red
    Write-Host "[!] The VHDX and folder have been prepared but the VM was not registered." -ForegroundColor Yellow
    Write-Host "[!] You can create the VM manually pointing to: $DestVHDX" -ForegroundColor Yellow
    exit 1
}

# --- Start the VM ---
Write-Host "[*] Starting VM..." -ForegroundColor Yellow
Start-VM -Name $VMName
Write-Host "[+] VM '$VMName' is starting." -ForegroundColor Green

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  Deployment Complete" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  VM Name    : $VMName"
Write-Host "  IP Address : $IPAddress"
Write-Host "  Domain     : $DomainName"
Write-Host ""
Write-Host "  The VM will boot, apply unattend.xml, set its hostname,"
Write-Host "  configure its static IP, and join the domain automatically."
Write-Host "  Allow a few minutes for first boot to complete."
Write-Host "============================================================`n" -ForegroundColor Cyan