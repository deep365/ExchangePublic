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

.PARAMETER Mode
    Selects which security context's work to perform. Office license state is
    split across the machine (SPP store, HKLM) and each user's profile (HKCU,
    LocalAppData, Credential Manager, WAM). Running everything as SYSTEM or as
    an elevated admin only cleans THAT identity's profile, leaving the real
    user's licenses untouched. Split the work instead:
      Machine - Admin/SYSTEM context. SPP license uninstall (WMI) + HKLM writes
                (User Settings reset keys, ClickToRun EmailAddress/TenantId/
                ProductKeys). Run from a SYSTEM scheduled task.
      User    - Logged-on user context (NO elevation needed). HKCU deletes,
                %LOCALAPPDATA% cache folders, Credential Manager, OneAuth,
                IdentityCache, WAM sign-out. Run from a scheduled task whose
                principal is the Interactive/logged-on user.
      Both    - Everything in one pass. Only correct when a single identity owns
                both the machine and the licensed profile (e.g. an interactive
                admin cleaning their own machine). Default.
      MachineThenUser - Run from a SYSTEM scheduled task. Performs the Machine
                half, then relaunches the User half inside the logged-on user's
                session via a transient scheduled task (waits, then cleans up).
                This is the recommended single-task deployment - see .NOTES.

.PARAMETER ClearO15
    Clear Office 2013 (15.0) licenses/keys. Default: $true

.PARAMETER ClearO16
    Clear Office 2016/365 (16.0) licenses/keys. Default: $true

.PARAMETER SignOutOfWAM
    Invoke SignOutOfWAMAccounts.ps1 if found in the same directory. Default: $true
    (User scope only.)

.PARAMETER SafeForRoamingUsers
    When $true the registry "Count" key is always set to 1, safe for roaming
    profiles. Set to $false only if the script may run more than once and you
    have no roaming profile users. Default: $true

.PARAMETER LogDir
    Custom folder for the log file. Defaults to %TEMP%.

.EXAMPLE
    .\OLicenseCleanup.ps1 -Verbose
    Run everything (Both) with full verbose output to the console.

.EXAMPLE
    .\OLicenseCleanup.ps1 -Mode Machine
    SPP uninstall + HKLM writes only. Run as SYSTEM/admin.

.EXAMPLE
    .\OLicenseCleanup.ps1 -Mode User
    Per-user cache cleanup only. Run as the logged-on (non-elevated) user.

.NOTES
    RECOMMENDED DEPLOYMENT - single SYSTEM scheduled task (Option 1):

      Trigger   : On user sign-in (or your custom event)
      Principal : SYSTEM  (Run with highest privileges)
      Action    : powershell.exe -NoProfile -ExecutionPolicy Bypass
                  -File "C:\Odm\OLicenseCleanup.ps1" -Mode MachineThenUser

    In MachineThenUser mode the SYSTEM process:
      1. Runs the Machine half itself (SPP uninstall + HKLM writes).
      2. Detects the active console user.
      3. Registers a TRANSIENT scheduled task that runs this same script with
         -Mode User as that interactive user (Limited / unelevated).
      4. Starts it, waits for completion, then deletes the transient task.

    This gives guaranteed ordering (machine completes before user half starts)
    with a single deployed task and no event-log handshake to maintain. Task
    Scheduler resolves the user's session token, so HKCU / %LOCALAPPDATA% /
    Credential Manager / WAM all resolve in the correct profile.

    ALTERNATIVE - two independently-triggered tasks (no self-relaunch):
      Task 1: SYSTEM principal,            -Mode Machine
      Task 2: Interactive 'Users' principal -Mode User
    Use this if you prefer fully decoupled tasks. Strict ordering is not
    actually required: the HKLM "User Settings\...\Delete" keys are consumed by
    Office on next app launch in the user context regardless of task order.

    RUN-ONCE BEHAVIOR (built in):
    The script writes completion markers so a login-triggered task only does
    real work once per scope, then becomes a fast no-op:
      Machine scope - marker at
        HKLM\SOFTWARE\Microsoft\Office\ScriptRun\OLicenseCleanup\CleanupCompleted
        Scope: once per machine, ever. Future SYSTEM runs skip the machine half.
      User scope    - marker at
        HKCU\...\ScriptRun\OLicenseCleanup\CleanupCompleted
        Scope: once per user, ever. Each user is cleaned on their first login;
        their later logins skip the user half. Other users are unaffected.
    Behavior is "always skip if already done" - there is no version re-trigger
    and no force override. To force a re-run, delete the relevant marker value.

    Original VBS author: Microsoft Customer Support Services
    PowerShell port:     Converted from v1.28
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    # Which half of the cleanup to run.
    #   Machine         - admin/SYSTEM context: SPP uninstall + HKLM writes.
    #   User            - logged-on user context: HKCU + LocalAppData + CredMan + WAM.
    #   Both            - everything in one pass (only correct when one identity
    #                     owns both, e.g. interactive admin = licensed user).
    #   MachineThenUser - run as SYSTEM: does the Machine half, then spawns the
    #                     User half inside the logged-on user's session via a
    #                     transient scheduled task, waits, and cleans up. This is
    #                     the mode to use for a SYSTEM-triggered scheduled task.
    [ValidateSet("Machine","User","Both","MachineThenUser")]
    [string] $Mode               = "Both",

    [string] $SkuFilter          = "O365",
    [bool]   $ClearO15           = $true,
    [bool]   $ClearO16           = $true,
    [bool]   $SignOutOfWAM        = $true,
    [bool]   $SafeForRoamingUsers = $true,
    [string] $LogDir             = "",

    # Internal: name used for the transient per-user relaunch task. Override only
    # if the default collides with an existing task in your environment.
    [string] $RelaunchTaskName   = "OLicenseCleanup_UserScope_Transient"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

#region -- Constants ------------------------------------------------------------
$SCRIPT_VERSION = "1.28"
$OFFICE_APP_ID  = "0ff1ce15-a989-479d-af46-f275c6370663"
$PS_WAM_SIGNOUT = "SignOutOfWAMAccounts.ps1"
$REG_LOG_PATH   = "HKCU:\Software\Microsoft\Office\16.0\Common\ScriptRun\OLicenseCleanup"

# Completion markers used to enforce run-once semantics:
#   Machine marker (HKLM) - once the machine half completes, future SYSTEM runs
#                           skip it. Scope: once per machine, ever.
#   User marker (HKCU)    - once a user's per-profile cleanup completes, that
#                           user's future logins skip it. Scope: once per user.
$REG_MACHINE_DONE = "HKLM:\SOFTWARE\Microsoft\Office\ScriptRun\OLicenseCleanup"
$REG_USER_DONE    = $REG_LOG_PATH   # reuse the existing HKCU key
$DONE_VALUE_NAME  = "CleanupCompleted"
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
$Script:RunMachine     = ($Mode -eq "Machine" -or $Mode -eq "Both" -or $Mode -eq "MachineThenUser")
$Script:RunUser        = ($Mode -eq "User"    -or $Mode -eq "Both")
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
#region -- Run-once completion guards -------------------------------------------
###############################################################################

#-------------------------------------------------------------------------------
#   Test-CleanupCompleted
#
#   Returns $true if the given scope has already completed a cleanup (marker
#   present and set to 1). Scope = "Machine" (HKLM) or "User" (HKCU).
#-------------------------------------------------------------------------------
function Test-CleanupCompleted {
    param([ValidateSet("Machine","User")][string]$Scope)

    $path = if ($Scope -eq "Machine") { $REG_MACHINE_DONE } else { $REG_USER_DONE }
    $val  = Get-RegDWord -Path $path -Name $DONE_VALUE_NAME
    return ($null -ne $val -and [int]$val -eq 1)
}

#-------------------------------------------------------------------------------
#   Set-CleanupCompleted
#
#   Writes the completion marker for the given scope, plus a timestamp and the
#   script version (informational only - presence of the marker is the gate).
#-------------------------------------------------------------------------------
function Set-CleanupCompleted {
    param([ValidateSet("Machine","User")][string]$Scope)

    $path = if ($Scope -eq "Machine") { $REG_MACHINE_DONE } else { $REG_USER_DONE }
    Set-RegDWord  -Path $path -Name $DONE_VALUE_NAME           -Value 1
    Set-RegString -Path $path -Name "CompletedTime"            -Value (Get-Date -Format "yyyyMMddHHmmss")
    Set-RegString -Path $path -Name "CompletedByVersion"       -Value $SCRIPT_VERSION
    Write-Log "$Scope-scope completion marker written to $path"
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

    # -- Run-attempt counter in registry (HKCU - only meaningful in User scope)
    if ($Script:RunUser) {
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
    else {
        Write-Verbose "=== Initialize: complete (Machine scope - skipped HKCU marker) ==="
    }
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
            Write-Verbose "Wow6432Node Office key found at '$key' --> 32-bit Office"
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
            Write-Verbose "Native Office key found at '$key' --> 64-bit Office"
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
                Write-Log "  --> Uninstall successful"
            }
            catch {
                Write-Log "  --> Uninstall failed: $_"
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

    # -- USER scope: direct delete of HKCU key -------------------------------
    if ($Script:RunUser) {
        $hkcuKey = "HKCU:\Software\Microsoft\Office\$Version.0\$RegKey"
        Write-Log "Remove key: $hkcuKey"
        try {
            Remove-Item -Path $hkcuKey -Recurse -Force -ErrorAction Stop
            $retVal = 0
            Write-Log "  --> Removed successfully"
        }
        catch {
            $retVal = 1
            Write-Log "  --> Key not found or removal failed (may be expected)"
        }

        $cacheLog = "LastRun" + ($RegKey -replace '\\', '') + "RegistryDelete"
        if ($Version -eq "16") {
            Set-RegDWord -Path $REG_LOG_PATH -Name $cacheLog -Value $retVal
        }
    }

    # -- MACHINE scope: create UserSettings key so Office resets on next launch
    if ($Script:RunMachine) {
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

        $strippedPattern = ($pattern -replace '\*', '')
        $patternStripped = [regex]::Escape($strippedPattern)

        foreach ($line in $lines) {
            $trimmed = $line.Trim()

            # Start collecting when we see the key name or OneDrive
            if (($trimmed -match [regex]::Escape($key)) -or ($trimmed -match "OneDrive")) {
                if ($trimmed -notmatch $patternStripped) {
                    $building   = $true
                    $targetLine = $trimmed
                    continue
                }
            }

            if ($building) {
                if ($trimmed -eq "") {
                    # Blank line = end of block --> submit
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
    Write-Log "  --> Return value: $retVal | $result"

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
            Write-Log "  --> Removed successfully"
        }
        catch {
            $retVal = 1
            Write-Log "  --> Failed to remove: $_"
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
        Write-Log "  --> Removed via Remove-Item"
    }
    catch {
        Write-Log "  --> Remove-Item failed: $_ - retrying with rd.exe"
    }

    # Fallback: rd /s /q
    if (Test-Path $FolderPath) {
        $retVal = (Start-Process -FilePath "cmd.exe" -ArgumentList "/c rd /s /q `"$FolderPath`"" `
                      -Wait -PassThru -WindowStyle Hidden).ExitCode
        Write-Log "  --> rd.exe exit code: $retVal"
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
        Write-Log "  --> Failed to launch WAM sign-out: $_"
    }

    Write-Log "  --> WAM sign-out return value: $retVal"
    Set-RegDWord -Path $REG_LOG_PATH -Name "LastRunLaunchSignOutOfWAM" -Value $retVal
}
#endregion

###############################################################################
#region -- User-scope relaunch (MachineThenUser) --------------------------------
###############################################################################

#-------------------------------------------------------------------------------
#   Get-ActiveConsoleUser
#
#   Returns the DOMAIN\User of the currently active console session, or $null
#   if no interactive user is logged on. Used by SYSTEM to know whose context
#   the user-scope cleanup must run in.
#-------------------------------------------------------------------------------
function Get-ActiveConsoleUser {
    # Win32_ComputerSystem.UserName reflects the console (interactive) user.
    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        if ($cs.UserName) {
            Write-Verbose "Active console user (Win32_ComputerSystem): $($cs.UserName)"
            return $cs.UserName
        }
    }
    catch {}

    # Fallback: parse `query user` for the Active session.
    try {
        $quser = & query user 2>$null
        foreach ($line in ($quser | Select-Object -Skip 1)) {
            if ($line -match '\bActive\b') {
                $name = ($line.TrimStart('>') -split '\s+')[0]
                if ($name) {
                    Write-Verbose "Active console user (query user): $name"
                    return $name
                }
            }
        }
    }
    catch {}

    return $null
}

#-------------------------------------------------------------------------------
#   Invoke-UserScopeRelaunch
#
#   Called when running as SYSTEM (-Mode MachineThenUser). Registers a transient
#   scheduled task that runs THIS script with -Mode User as the logged-on user,
#   runs it, waits for completion, then deletes it. Task Scheduler handles the
#   session/token so the per-user caches resolve in the correct profile.
#-------------------------------------------------------------------------------
function Invoke-UserScopeRelaunch {
    Write-LogSubHeader "Relaunching user-scope cleanup in the logged-on user's session"

    $targetUser = Get-ActiveConsoleUser
    if (-not $targetUser) {
        Write-Log "No interactive user is logged on - skipping user-scope relaunch."
        return
    }
    Write-Log "Target interactive user: $targetUser"

    # Resolve this script's own full path to re-invoke it.
    $scriptPath = $PSCommandPath
    if (-not $scriptPath) { $scriptPath = $MyInvocation.MyCommand.Path }
    if (-not (Test-Path $scriptPath)) {
        Write-Log "ERROR: Could not resolve own script path ('$scriptPath') - cannot relaunch."
        return
    }

    # Build the argument string for the user-scope pass. Forward the same
    # behavioral switches; force -Mode User so the relaunched copy does only
    # the per-user work.
    $argLine = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`" " +
               "-Mode User " +
               "-SkuFilter `"$SkuFilter`" " +
               "-ClearO15 `$$ClearO15 -ClearO16 `$$ClearO16 " +
               "-SignOutOfWAM `$$SignOutOfWAM -SafeForRoamingUsers `$$SafeForRoamingUsers"
    if ($LogDir -ne "") { $argLine += " -LogDir `"$LogDir`"" }

    Write-Verbose "Relaunch command: powershell.exe $argLine"

    # Clean up any stale instance of the transient task first.
    try {
        Unregister-ScheduledTask -TaskName $RelaunchTaskName -Confirm:$false -ErrorAction SilentlyContinue
    }
    catch {}

    try {
        $action    = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $argLine
        $principal = New-ScheduledTaskPrincipal -UserId $targetUser -LogonType Interactive -RunLevel Limited
        $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries `
                        -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 15)

        Register-ScheduledTask -TaskName $RelaunchTaskName -Action $action `
            -Principal $principal -Settings $settings -Force -ErrorAction Stop | Out-Null
        Write-Log "Registered transient task '$RelaunchTaskName' for $targetUser"

        Start-ScheduledTask -TaskName $RelaunchTaskName -ErrorAction Stop
        Write-Log "Started transient user-scope task; waiting for completion..."

        # Wait for the task to start then finish (cap the wait).
        $deadline = (Get-Date).AddMinutes(15)
        Start-Sleep -Seconds 2
        do {
            Start-Sleep -Seconds 2
            $info  = Get-ScheduledTask -TaskName $RelaunchTaskName -ErrorAction SilentlyContinue
            $state = if ($info) { $info.State } else { "Unknown" }
        } while ($state -eq "Running" -and (Get-Date) -lt $deadline)

        $taskInfo = Get-ScheduledTaskInfo -TaskName $RelaunchTaskName -ErrorAction SilentlyContinue
        if ($taskInfo) {
            Write-Log "User-scope task finished. LastTaskResult: $($taskInfo.LastTaskResult)"
        }
        else {
            Write-Log "User-scope task finished (no result info available)."
        }
    }
    catch {
        Write-Log "ERROR during user-scope relaunch: $_"
    }
    finally {
        # Always remove the transient task.
        try {
            Unregister-ScheduledTask -TaskName $RelaunchTaskName -Confirm:$false -ErrorAction SilentlyContinue
            Write-Verbose "Removed transient task '$RelaunchTaskName'"
        }
        catch {
            Write-Log "WARN: Could not remove transient task '$RelaunchTaskName': $_"
        }
    }
}
#endregion

###############################################################################
#region -- MAIN -----------------------------------------------------------------
###############################################################################

try {
    Initialize

    Write-LogH2 "Cleanup start (Mode: $Mode | Machine=$($Script:RunMachine) User=$($Script:RunUser))"

    # ---- Run-once guards: skip a scope that has already completed ----------
    # Machine scope: once per machine, ever (HKLM marker).
    # User scope:    once per user, ever (HKCU marker, this profile only).
    $doMachine = $Script:RunMachine
    $doUser    = $Script:RunUser

    if ($doMachine -and (Test-CleanupCompleted -Scope "Machine")) {
        Write-Log "Machine-scope cleanup already completed on this machine - skipping."
        $doMachine = $false
    }
    if ($doUser -and (Test-CleanupCompleted -Scope "User")) {
        Write-Log "User-scope cleanup already completed for this user - skipping."
        $doUser = $false
    }

    # Make the split-context Reset-UserKeyEx honor the post-guard decisions.
    $Script:RunMachine = $doMachine
    $Script:RunUser    = $doUser

    # ---- MACHINE scope (admin / SYSTEM): SPP license uninstall + HKLM ------
    if ($doMachine) {
        Invoke-CleanOSPP     -Filter $SkuFilter
        Clear-ConfigUser     # HKLM ClickToRun EmailAddress / TenantId / ProductKeys
    }

    # ---- Reset-UserKey: internally split (HKCU delete = User scope, --------
    #      HKLM User Settings write = Machine scope). Each half is gated by
    #      $Script:RunMachine / $Script:RunUser set above. ------------------
    if ($doMachine -or $doUser) {
        Reset-UserKey        -RegKey "Common\Identity"                -CustomName ""
        Reset-UserKey        -RegKey "Common\Roaming\Identities"      -CustomName ""
        Reset-UserKey        -RegKey "Common\Internet\WebServiceCache" -CustomName ""
        Reset-UserKey        -RegKey "Common\ServicesManagerCache"    -CustomName ""
        Reset-UserKey        -RegKey "Common\Licensing"               -CustomName ""
        Reset-UserKey        -RegKey "Registration"                   -CustomName ""
    }

    # ---- USER scope (logged-on user): per-profile caches ------------------
    if ($doUser) {
        Clear-CredmanCache
        Clear-SCALicCache
        Clear-VNextLicCache
        Clear-OneAuthCache
        Clear-IdentityCache

        if ($SignOutOfWAM) { Invoke-SignOutOfWAM }
    }

    # ---- Write completion markers for whatever ran successfully -----------
    if ($doMachine) { Set-CleanupCompleted -Scope "Machine" }
    if ($doUser)    { Set-CleanupCompleted -Scope "User" }

    # ---- MachineThenUser: relaunch the user half in the user's session ----
    # Always attempt the relaunch (even if THIS machine half was skipped):
    # a different user may not yet be cleaned. The relaunched -Mode User copy
    # self-guards on its own HKCU marker and exits fast if already done.
    if ($Mode -eq "MachineThenUser") {
        Invoke-UserScopeRelaunch
    }

    Write-LogH2 "Cleanup end"
    Write-Host "`nOffice license cleanup completed (Mode: $Mode). Log: $($Script:LogFilePath)" -ForegroundColor Green
}
finally {
    if ($Script:LogStream) {
        $Script:LogStream.Flush()
        $Script:LogStream.Close()
    }
}
#endregion
