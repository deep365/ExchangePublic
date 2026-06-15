#Requires -Version 5.1
<#
.SYNOPSIS
    OLicenseCleanup.ps1 - v1.28
    Microsoft Customer Support Services - Office License Reset Utility

.DESCRIPTION
    Removes all licenses for Office 2013 and 2016 from the (Office) Software
    Protection Platform, clears credential caches, registry identity keys, and
    optional WAM sign-out.

.PARAMETER Verbose
    Switch that enables verbose console output for every action taken.
    PowerShell's built-in -Verbose switch is used; pass it at invocation:
        .\OLicenseCleanup.ps1 -Verbose

.PARAMETER SkuFilter
    Controls which SPP licenses are removed.
      "O365"    - (default) Remove licenses whose name contains "O365"
      ""        - Remove ALL Office licenses
      "NOTO365" - Remove licenses that do NOT contain "O365"

.PARAMETER ClearO15
    Clear Office 2013 (15.0) licenses/keys. Default: $true

.PARAMETER ClearO16
    Clear Office 2016/365 (16.0) licenses/keys. Default: $true

.PARAMETER SignOutOfWAM
    Invoke SignOutOfWAMAccounts.ps1 if found in the same directory. Default: $true

.PARAMETER SafeForRoamingUsers
    When $true the registry "Count" key is always set to 1, safe for roaming
    profiles. Set to $false only if the script may run more than once and you
    have no roaming profile users. Default: $true

.PARAMETER LogDir
    Custom folder for the log file. Defaults to %TEMP%.

.EXAMPLE
    .\OLicenseCleanup.ps1 -Verbose
    Run with full verbose output to the console.

.EXAMPLE
    .\OLicenseCleanup.ps1 -SkuFilter "" -Verbose
    Remove ALL Office SPP licenses with verbose output.

.NOTES
    Must be run as an Administrator for HKLM registry writes and WMI SPP access.
    Original VBS author: Microsoft Customer Support Services
    PowerShell port:     Converted from v1.28
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $SkuFilter          = "O365",
    [bool]   $ClearO15           = $false,
    [bool]   $ClearO16           = $false,
    [bool]   $SignOutOfWAM        = $true,
    [bool]   $SafeForRoamingUsers = $true,
    [string] $LogDir             = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

#region -- Constants ------------------------------------------------------------
$SCRIPT_VERSION = "1.28"
$OFFICE_APP_ID  = "0ff1ce15-a989-479d-af46-f275c6370663"
$PS_WAM_SIGNOUT = "SignOutOfWAMAccounts.ps1"
$REG_LOG_PATH   = "HKCU:\Software\Microsoft\Office\16.0\Common\ScriptRun\OLicenseCleanup"
#endregion

#region -- Script-scope state ---------------------------------------------------
$Script:LogStream      = $null
$Script:LogFilePath    = $null
$Script:Is64BitOS      = $false
$Script:Is64BitOffice  = $false
$Script:OSInfo         = ""
$Script:ProfilesDir    = ""
$Script:ScriptDir      = Split-Path -Parent $MyInvocation.MyCommand.Path
$Script:TimeStamp      = ""
#endregion

###############################################################################
#region -- Logging helpers ------------------------------------------------------
###############################################################################

function Write-LogRaw {
    param([string]$Line)
    if ($Script:LogStream) { $Script:LogStream.WriteLine($Line) }
}

# LogH  - major section header (==== underline)
function Write-LogHeader {
    param([string]$Message)
    $underline = "=" * $Message.Length
    Write-LogRaw ""
    Write-LogRaw "$Message`r`n$underline"
    Write-Verbose ""
    Write-Verbose "$Message"
    Write-Verbose $underline
}

# LogH1 - sub-section header (---- underline)
function Write-LogSubHeader {
    param([string]$Message)
    $underline = "-" * $Message.Length
    Write-LogRaw ""
    Write-LogRaw "$Message`r`n$underline"
    Write-Verbose ""
    Write-Verbose $Message
    Write-Verbose $underline
}

# LogH2 - plain header, no underline
function Write-LogH2 {
    param([string]$Message)
    Write-LogRaw ""
    Write-LogRaw $Message
    Write-Verbose $Message
}

# Log   - timestamped entry, echoed to console
function Write-Log {
    param([string]$Message = "")
    if ($Message -eq "") {
        Write-LogRaw ""
    }
    else {
        $entry = "   $(Get-Date -Format 'HH:mm:ss'): $Message"
        Write-LogRaw $entry
        Write-Verbose $entry
    }
}

# LogOnly - timestamped entry written only to the log file
function Write-LogOnly {
    param([string]$Message = "")
    if ($Message -eq "") {
        Write-LogRaw ""
    }
    else {
        Write-LogRaw "   $(Get-Date -Format 'HH:mm:ss'): $Message"
    }
}
#endregion

###############################################################################
#region -- Registry helpers -----------------------------------------------------
###############################################################################

function Get-RegDWord {
    param([string]$Path, [string]$Name)
    try {
        $val = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
        return $val.$Name
    }
    catch { return $null }
}

function Set-RegDWord {
    param([string]$Path, [string]$Name, [int]$Value)
    try {
        if (-not (Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }
        Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type DWord -ErrorAction Stop
    }
    catch { Write-Log "WARN: Could not set DWORD '$Name' at '$Path': $_" }
}

function Set-RegString {
    param([string]$Path, [string]$Name, [string]$Value)
    try {
        if (-not (Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }
        Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type String -ErrorAction Stop
    }
    catch { Write-Log "WARN: Could not set String '$Name' at '$Path': $_" }
}

function Remove-RegValue {
    param([string]$Path, [string]$Name)
    try {
        Remove-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
    }
    catch {}   # silently ignore if not present
}

# Convert HKLM/HKCU WMI-style path to PS provider path
function ConvertTo-PSRegPath {
    param([string]$Path)
    $Path = $Path -replace '^HKLM\\', 'HKLM:\'
    $Path = $Path -replace '^HKCU\\', 'HKCU:\'
    return $Path
}
#endregion

###############################################################################
#region -- Initialize -----------------------------------------------------------
###############################################################################

function Initialize {
    Write-Verbose "=== Initialize: starting ==="

    # -- OS bitness ----------------------------------------------------------
    $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
    if ($cs) {
        $Script:Is64BitOS = $cs.SystemType -match "^x64"
        Write-Verbose "SystemType reported: $($cs.SystemType)"
    }
    Write-Verbose "64-bit OS detected: $($Script:Is64BitOS)"

    # -- Office bitness ------------------------------------------------------
    $Script:Is64BitOffice = Detect-OfficeBitness
    Write-Verbose "64-bit Office detected: $($Script:Is64BitOffice)"

    # -- OS info -------------------------------------------------------------
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
    if ($os) {
        $Script:OSInfo = "$($os.Caption)$($os.OtherTypeDescription), SP $($os.ServicePackMajorVersion), " +
                         "Version: $($os.Version), Codepage: $($os.CodeSet), " +
                         "Country Code: $($os.CountryCode), Language: $($os.OSLanguage)"
    }
    Write-Verbose "OS Info: $($Script:OSInfo)"

    # -- Log directory / file ------------------------------------------------
    if ($LogDir -eq "") { $LogDir = $env:TEMP }
    $ts           = Get-Date -Format "yyyyMMddHHmmss"
    $Script:TimeStamp = $ts
    $logName      = "$LogDir\$($env:COMPUTERNAME)_${ts}_OLicenseClean.txt"
    try {
        $Script:LogStream   = [System.IO.StreamWriter]::new($logName, $false, [System.Text.Encoding]::Unicode)
        $Script:LogFilePath = $logName
    }
    catch {
        # Fallback to %TEMP%
        $logName = "$($env:TEMP)\$($env:COMPUTERNAME)_${ts}_OLicenseClean.txt"
        $Script:LogStream   = [System.IO.StreamWriter]::new($logName, $false, [System.Text.Encoding]::Unicode)
        $Script:LogFilePath = $logName
    }
    Write-Verbose "Log file: $($Script:LogFilePath)"

    Write-LogH2 ("Microsoft Customer Support Services - Office License Reset Utility`r`n`r`n" +
                 "Version:`t$SCRIPT_VERSION`r`n" +
                 "64-bit OS:`t$($Script:Is64BitOS)`r`n" +
                 "64-bit Office:`t$($Script:Is64BitOffice)`r`n" +
                 "Cleanup start:`t$(Get-Date -Format 'HH:mm:ss')")
    Write-LogH2 "OS Details: $($Script:OSInfo)`r`n"

    Write-LogOnly "Remove O15 Lic: $ClearO15"
    Write-LogOnly "Remove O16 Lic: $ClearO16"
    Write-LogOnly "Verbose mode:   $($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Verbose'))"

    # -- Profiles directory --------------------------------------------------
    try {
        $Script:ProfilesDir = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" `
                                    -Name ProfilesDirectory -ErrorAction Stop).ProfilesDirectory
        $Script:ProfilesDir = [System.Environment]::ExpandEnvironmentVariables($Script:ProfilesDir)
    }
    catch {}
    if (-not (Test-Path $Script:ProfilesDir)) {
        $Script:ProfilesDir = Split-Path $env:USERPROFILE -Parent
    }
    Write-LogOnly "Users profile location: $($Script:ProfilesDir)"
    Write-LogOnly "Current Directory: $($Script:ScriptDir)"

    # -- Run-attempt counter in registry ------------------------------------
    $attempts = Get-RegDWord -Path $REG_LOG_PATH -Name "ScriptRunAttempts"
    if ($null -ne $attempts) { $attempts = [int]$attempts + 1 } else { $attempts = 1 }

    # Clear old key and recreate
    try { Remove-Item -Path $REG_LOG_PATH -Recurse -Force -ErrorAction SilentlyContinue } catch {}

    New-Item -Path $REG_LOG_PATH -Force | Out-Null
    Set-RegString -Path $REG_LOG_PATH -Name "Version"          -Value $SCRIPT_VERSION
    Set-RegString -Path $REG_LOG_PATH -Name "LastRunTime"       -Value $Script:TimeStamp
    Set-RegString -Path $REG_LOG_PATH -Name "LastRunDirectory"  -Value $Script:ScriptDir
    Set-RegDWord  -Path $REG_LOG_PATH -Name "ScriptRunAttempts" -Value $attempts
    Set-RegDWord  -Path $REG_LOG_PATH -Name "SafeForRoamingUsers" -Value ([int]$SafeForRoamingUsers)

    Write-Verbose "=== Initialize: complete (attempt #$attempts) ==="
}
#endregion

###############################################################################
#region -- Detect OS / Office bitness ------------------------------------------
###############################################################################

function Detect-OfficeBitness {
    if (-not $Script:Is64BitOS) { return $false }

    $checks = @(
        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
        "HKLM:\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration"
    )
    foreach ($key in $checks) {
        try {
            $p = (Get-ItemProperty $key -Name "platform" -ErrorAction Stop).platform
            Write-Verbose "Office platform from '$key': $p"
            return ($p -eq "x64")
        }
        catch {}
    }

    $bagChecks = @(
        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\propertyBag",
        "HKLM:\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag"
    )
    foreach ($key in $bagChecks) {
        try {
            $p = (Get-ItemProperty $key -Name "Platform" -ErrorAction Stop).Platform
            Write-Verbose "Office platform from propertyBag '$key': $p"
            return ($p -eq "x64")
        }
        catch {}
    }

    $wow64Checks = @(
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\Common\InstallRoot",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Common\InstallRoot"
    )
    foreach ($key in $wow64Checks) {
        try {
            Get-ItemProperty $key -Name "Path" -ErrorAction Stop | Out-Null
            Write-Verbose "Wow6432Node Office key found at '$key' -> 32-bit Office"
            return $false
        }
        catch {}
    }

    $nativeChecks = @(
        "HKLM:\SOFTWARE\Microsoft\Office\Common\InstallRoot",
        "HKLM:\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot"
    )
    foreach ($key in $nativeChecks) {
        try {
            Get-ItemProperty $key -Name "Path" -ErrorAction Stop | Out-Null
            Write-Verbose "Native Office key found at '$key' -> 64-bit Office"
            return $true
        }
        catch {}
    }

    return $false
}

function Get-WindowsVersionNT {
    $os  = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
    $ver = $os.Version -split '\.'
    return ([int]$ver[0] * 100 + [int]$ver[1])
}
#endregion

###############################################################################
#region -- CleanOSPP ------------------------------------------------------------
###############################################################################

function Invoke-CleanOSPP {
    param([string]$Filter)

    Write-LogSubHeader "Cleaning Office SPP licenses (filter: '$Filter')"

    $winNT = Get-WindowsVersionNT
    Write-Verbose "Windows NT version number: $winNT"

    $wmiClass = if ($winNT -gt 601) { "SoftwareLicensingProduct" } else { "OfficeSoftwareProtectionProduct" }
    Write-Verbose "Using WMI class: $wmiClass"

    $query = "SELECT ID, ApplicationId, PartialProductKey, Description, Name, ProductKeyID FROM $wmiClass " +
             "WHERE ApplicationId = '$OFFICE_APP_ID' AND PartialProductKey <> NULL"

    $products = Get-CimInstance -Query $query -ErrorAction SilentlyContinue
    if (-not $products) {
        Write-Log "No Office SPP licenses found."
        return
    }

    foreach ($pi in $products) {
        Write-Log "License found: $($pi.Name)"
        if ($null -eq $pi) { continue }

        $shouldRemove = $false

        if ($Filter -eq "" -or $pi.Name -match [regex]::Escape($Filter)) {
            $shouldRemove = $true
        }
        if ($Filter.ToUpper() -eq "NOTO365" -and $pi.Name -notmatch "O365") {
            $shouldRemove = $true
        }

        if ($shouldRemove) {
            Write-Log "Uninstalling ProductKey: $($pi.Name) - Key: $($pi.ProductKeyID)"
            try {
                Invoke-CimMethod -InputObject $pi -MethodName "UninstallProductKey" `
                    -Arguments @{ ProductKey = $pi.ProductKeyID } -ErrorAction Stop | Out-Null
                Write-Log "  -> Uninstall successful"
            }
            catch {
                Write-Log "  -> Uninstall failed: $_"
            }
            $safeKeyName = $pi.Name -replace '[^A-Za-z0-9]', '_'
            Set-RegDWord -Path $REG_LOG_PATH -Name "LastRunUninstallSPPkey_$safeKeyName" -Value 0
        }
    }
}
#endregion

###############################################################################
#region -- ResetUserKey ---------------------------------------------------------
###############################################################################

function Reset-UserKey {
    param([string]$RegKey, [string]$CustomName)

    Write-Verbose "Reset-UserKey: '$RegKey'"
    if ($ClearO15) { Reset-UserKeyEx -RegKey $RegKey -CustomName $CustomName -Version "15" }
    if ($ClearO16) { Reset-UserKeyEx -RegKey $RegKey -CustomName $CustomName -Version "16" }
}

function Reset-UserKeyEx {
    param([string]$RegKey, [string]$CustomName, [string]$Version)

    if ($CustomName -eq "") { $CustomName = "CustomUserReset" }

    # -- Direct delete of HKCU key -------------------------------------------
    $hkcuKey = "HKCU:\Software\Microsoft\Office\$Version.0\$RegKey"
    Write-Log "Remove key: $hkcuKey"
    try {
        Remove-Item -Path $hkcuKey -Recurse -Force -ErrorAction Stop
        $retVal = 0
        Write-Log "  -> Removed successfully"
    }
    catch {
        $retVal = 1
        Write-Log "  -> Key not found or removal failed (may be expected)"
    }

    $cacheLog = "LastRun" + ($RegKey -replace '\\', '') + "RegistryDelete"
    if ($Version -eq "16") {
        Set-RegDWord -Path $REG_LOG_PATH -Name $cacheLog -Value $retVal
    }

    # -- Create UserSettings key so Office resets it on next launch ----------
    if ($Script:Is64BitOS -and $Script:Is64BitOffice) {
        $settingsKey = "HKLM:\SOFTWARE\Microsoft\Office\$Version.0\User Settings"
    }
    elseif (-not $Script:Is64BitOS) {
        $settingsKey = "HKLM:\SOFTWARE\Microsoft\Office\$Version.0\User Settings"
    }
    else {
        $settingsKey = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\$Version.0\User Settings"
    }

    $squirrelPath = "$CustomName\Delete\Software\Microsoft\Office\$Version.0\$RegKey"
    $fullPath     = "$settingsKey\$squirrelPath"

    Write-Verbose "Creating Office UserSettings reset key: $fullPath"
    try {
        New-Item -Path $fullPath -Force -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Log "WARN: Could not create settings key '$fullPath': $_"
    }

    $countKey = "$settingsKey\$CustomName"
    $count    = 1
    if (-not $SafeForRoamingUsers) {
        $existing = Get-RegDWord -Path $countKey -Name "Count"
        if ($null -ne $existing) { $count = [int]$existing + 1 }
    }

    Set-RegDWord -Path $countKey -Name "Count" -Value $count
    Set-RegDWord -Path $countKey -Name "Order"  -Value 1
    Write-LogOnly "Add SettingsKey: $fullPath"
    Write-LogOnly "Count: $count"
}
#endregion

###############################################################################
#region -- ClearCredmanCache ----------------------------------------------------
###############################################################################

function Clear-CredmanCache {
    Write-LogSubHeader "Clearing Credential Manager cache"

    $credPatterns = @{
        "MicrosoftOffice1"  = "MicrosoftOffice1*"
        "msteams"           = "msteams*"
        "Microsoft_OC"      = "Microsoft_OC*"
        "OneDrive Cached"   = "OneDrive*"
    }

    foreach ($key in $credPatterns.Keys) {
        $pattern = $credPatterns[$key]
        Write-Verbose "Querying credentials matching: $pattern"

        $output = & cmdkey.exe /list:$pattern 2>&1 | Out-String

        # Parse each target line from cmdkey output
        $lines      = $output -split '\r?\n'
        $building   = $false
        $targetLine = ""

        foreach ($line in $lines) {
            $trimmed = $line.Trim()

            # Start collecting when we see the key name or OneDrive
            if (($trimmed -match [regex]::Escape($key)) -or ($trimmed -match "OneDrive")) {
                if ($trimmed -notmatch [regex]::Escape($pattern -replace '\*', '')) {
                    $building   = $true
                    $targetLine = $trimmed
                    continue
                }
            }

            if ($building) {
                if ($trimmed -eq "") {
                    # Blank line = end of block -> submit
                    if ($targetLine -ne "") {
                        Remove-CredmanEntry $targetLine
                    }
                    $building   = $false
                    $targetLine = ""
                }
                else {
                    $targetLine += " $trimmed"
                }
            }
        }

        # Flush any trailing entry
        if ($building -and $targetLine -ne "") {
            Remove-CredmanEntry $targetLine
        }
    }
}

function Remove-CredmanEntry {
    param([string]$TargetLine)

    $TargetLine = $TargetLine.Trim()
    Write-Log "Remove from CredmanCache: $TargetLine"
    Write-Log "Execute removal: cmdkey.exe /delete:`"$TargetLine`""

    $result = & cmdkey.exe /delete:"$TargetLine" 2>&1
    $retVal = $LASTEXITCODE
    Write-Log "  -> Return value: $retVal | $result"

    Set-RegDWord -Path $REG_LOG_PATH -Name "LastRunCmdKeyDelete" -Value $retVal
}
#endregion

###############################################################################
#region -- ClearSCALicCache -----------------------------------------------------
###############################################################################

function Clear-SCALicCache {
    Write-LogSubHeader "Clearing SCA (SharedComputerActivation) license cache"

    $localAppData = $env:LOCALAPPDATA

    if ($ClearO15) {
        Remove-CachedFolder -FolderPath "$localAppData\Microsoft\Office\15.0\Licensing" `
                            -CacheName "LocalAppDataO15LicensingFolderDelete"
    }
    if ($ClearO16) {
        Remove-CachedFolder -FolderPath "$localAppData\Microsoft\Office\16.0\Licensing" `
                            -CacheName "LocalAppDataLicensingFolderDelete"
    }

    # Custom token location override
    try {
        $override = (Get-ItemProperty "HKLM:\Software\Microsoft\Office\16.0\Common\Licensing" `
                        -Name "SCLCacheOverrideDirectory" -ErrorAction Stop).SCLCacheOverrideDirectory
        if ($override) {
            Write-Verbose "Custom SCA cache override directory: $override"
            Remove-CachedFolder -FolderPath $override -CacheName "SCLCacheOverrideDirectory"
        }
    }
    catch {}
}
#endregion

###############################################################################
#region -- ClearVNextLicCache ---------------------------------------------------
###############################################################################

function Clear-VNextLicCache {
    Write-LogSubHeader "Clearing vNext license cache"

    $localAppData   = $env:LOCALAPPDATA
    $programData    = $env:PROGRAMDATA

    Remove-CachedFolder -FolderPath "$localAppData\Microsoft\Office\Licenses" `
                        -CacheName "LocalAppDataLicensesFolderDelete"

    # Device Based Licensing cache (note: VBS used %localappdata% here - preserved)
    Remove-CachedFolder -FolderPath "$localAppData\Microsoft\Licenses" `
                        -CacheName "ProgramDataLicensesFolderDelete"
}
#endregion

###############################################################################
#region -- ClearIdentityCache ---------------------------------------------------
###############################################################################

function Clear-IdentityCache {
    Write-LogSubHeader "Clearing Windows Identity cache"
    $localAppData = $env:LOCALAPPDATA
    Remove-CachedFolder -FolderPath "$localAppData\Microsoft\IdentityCache" `
                        -CacheName "LocalAppDataIdentityCacheFolderDelete"
}
#endregion

###############################################################################
#region -- ClearOneAuthCache ----------------------------------------------------
###############################################################################

function Clear-OneAuthCache {
    Write-LogSubHeader "Clearing OneAuth cache"
    $localAppData = $env:LOCALAPPDATA
    Remove-CachedFolder -FolderPath "$localAppData\Microsoft\OneAuth" `
                        -CacheName "LocalAppDataOneAuthFolderDelete"
}
#endregion

###############################################################################
#region -- ClearConfigUser ------------------------------------------------------
###############################################################################

function Clear-ConfigUser {
    Write-LogSubHeader "Clearing HKLM cached user identity values (EmailAddress / TenantId / ProductKeys)"

    if (-not $ClearO16) {
        Write-Verbose "ClearO16 is false - skipping ClearConfigUser"
        return
    }

    $configKey = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"

    try {
        $props = Get-ItemProperty -Path $configKey -ErrorAction Stop
    }
    catch {
        Write-Log "ClickToRun Configuration key not found - skipping"
        return
    }

    $props.PSObject.Properties | Where-Object {
        $_.Name -match '\.EmailAddress$' -or
        $_.Name -match '\.TenantId$'     -or
        $_.Name -eq 'ProductKeys'
    } | ForEach-Object {
        $valueName = $_.Name
        Write-Log "Remove entry: HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration\$valueName"
        try {
            Remove-ItemProperty -Path $configKey -Name $valueName -ErrorAction Stop
            $retVal = 0
            Write-Log "  -> Removed successfully"
        }
        catch {
            $retVal = 1
            Write-Log "  -> Failed to remove: $_"
        }
        $safeVal = $valueName -replace '[^A-Za-z0-9]', '_'
        Set-RegDWord -Path $REG_LOG_PATH -Name "LastRun${safeVal}RegistryDelete" -Value $retVal
    }
}
#endregion

###############################################################################
#region -- ClearFolder / Remove-CachedFolder ------------------------------------
###############################################################################

function Remove-CachedFolder {
    param([string]$FolderPath, [string]$CacheName)

    $regName = "LastRun$CacheName"

    if (-not (Test-Path $FolderPath)) {
        Write-Log "Folder not found (skipping): $FolderPath"
        Set-RegDWord -Path $REG_LOG_PATH -Name $regName -Value 3
        return
    }

    Write-Log "Remove folder: $FolderPath"

    # Strip read-only attribute if set
    try {
        $dir = Get-Item $FolderPath -ErrorAction Stop
        if ($dir.Attributes -band [System.IO.FileAttributes]::ReadOnly) {
            $dir.Attributes = $dir.Attributes -bxor [System.IO.FileAttributes]::ReadOnly
            Write-Verbose "  Cleared ReadOnly attribute on: $FolderPath"
        }
    }
    catch {}

    # Attempt PowerShell removal first
    try {
        Remove-Item -Path $FolderPath -Recurse -Force -ErrorAction Stop
        Write-Log "  -> Removed via Remove-Item"
    }
    catch {
        Write-Log "  -> Remove-Item failed: $_ - retrying with rd.exe"
    }

    # Fallback: rd /s /q
    if (Test-Path $FolderPath) {
        $retVal = (Start-Process -FilePath "cmd.exe" -ArgumentList "/c rd /s /q `"$FolderPath`"" `
                      -Wait -PassThru -WindowStyle Hidden).ExitCode
        Write-Log "  -> rd.exe exit code: $retVal"
        Set-RegDWord -Path $REG_LOG_PATH -Name $regName -Value $retVal
    }
    else {
        Set-RegDWord -Path $REG_LOG_PATH -Name $regName -Value 0
    }
}
#endregion

###############################################################################
#region -- InvokeSignOutOfWAM ---------------------------------------------------
###############################################################################

function Invoke-SignOutOfWAM {
    Write-LogSubHeader "Invoking WAM sign-out script"

    $wamScript = Join-Path $Script:ScriptDir $PS_WAM_SIGNOUT
    if (-not (Test-Path $wamScript)) {
        Write-Log "$PS_WAM_SIGNOUT not found in script directory - skipping WAM sign-out"
        Set-RegDWord -Path $REG_LOG_PATH -Name "LastRunLaunchSignOutOfWAM" -Value 2
        return
    }

    Write-Log "Running: powershell.exe -ExecutionPolicy Unrestricted -NoProfile -File `"$wamScript`""
    try {
        $proc = Start-Process -FilePath "powershell.exe" `
            -ArgumentList "-ExecutionPolicy Unrestricted -NoProfile -WindowStyle Hidden -File `"$wamScript`"" `
            -Wait -PassThru -WindowStyle Hidden -ErrorAction Stop
        $retVal = $proc.ExitCode
    }
    catch {
        $retVal = -1
        Write-Log "  -> Failed to launch WAM sign-out: $_"
    }

    Write-Log "  -> WAM sign-out return value: $retVal"
    Set-RegDWord -Path $REG_LOG_PATH -Name "LastRunLaunchSignOutOfWAM" -Value $retVal
}
#endregion

###############################################################################
#region -- MAIN -----------------------------------------------------------------
###############################################################################

try {
    Initialize

    Write-LogH2 "Cleanup start"

    Invoke-CleanOSPP         -Filter $SkuFilter

    Reset-UserKey            -RegKey "Common\Identity"            -CustomName ""
    Reset-UserKey            -RegKey "Common\Roaming\Identities"  -CustomName ""
    Reset-UserKey            -RegKey "Common\Internet\WebServiceCache" -CustomName ""
    Reset-UserKey            -RegKey "Common\ServicesManagerCache" -CustomName ""
    Reset-UserKey            -RegKey "Common\Licensing"           -CustomName ""
    Reset-UserKey            -RegKey "Registration"               -CustomName ""

    Clear-CredmanCache
    Clear-SCALicCache
    Clear-ConfigUser
    Clear-VNextLicCache
    Clear-OneAuthCache
    Clear-IdentityCache

    if ($SignOutOfWAM) { Invoke-SignOutOfWAM }

    Write-LogH2 "Cleanup end"
    Write-Host "`nOffice license cleanup completed. Log: $($Script:LogFilePath)" -ForegroundColor Green
}
finally {
    if ($Script:LogStream) {
        $Script:LogStream.Flush()
        $Script:LogStream.Close()
    }
}
#endregion
