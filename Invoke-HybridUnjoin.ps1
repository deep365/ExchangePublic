#Requires -Version 5.1
<#
.SYNOPSIS
    Invoke-HybridUnjoin.ps1 - v2.0
    Prerequisite for OLicenseCleanup.ps1 in a SOURCE->TARGET tenant migration.

.DESCRIPTION
    Unjoins a Windows device from the SOURCE Microsoft Entra ID tenant ONLY IF
    the device is currently registered to that source tenant, then optionally
    writes the TARGET tenant's CDJ (Cloud Domain Join) registry keys so the
    device hybrid-joins the destination tenant on the next Automatic-Device-Join
    run.

    Guard: the script reads 'dsregcmd /status', extracts the current TenantId,
    and proceeds ONLY when it matches -SourceTenantId. For any other state
    (already on the target tenant, unjoined, or a different tenant) it logs and
    exits without changes. This makes it safe to run across a mixed fleet.

    The source tenant SCP has already been removed in this environment, so the
    Automatic-Device-Join scheduled task can no longer pull the device back to
    the source. This script therefore does NOT disable that task; instead it
    writes the TARGET CDJ keys to steer the (re)join to the destination tenant.

    Run order for a migrating device:
      1. Invoke-HybridUnjoin.ps1   (this script - elevated admin)
      2. Reboot
      3. OLicenseCleanup.ps1       (clears the now-unseeded Office caches)
      4. Open Office and activate with the new (target) account

.PARAMETER SourceTenantId
    GUID of the SOURCE tenant. The unjoin only runs if the device's current
    TenantId matches this value. Default: the source tenant for this migration.

.PARAMETER TargetTenantId
    GUID of the TARGET tenant, written to
    HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\CDJ\AAD\TenantId
    Default: the target tenant for this migration.

.PARAMETER TargetTenantName
    Verified domain name of the TARGET tenant (e.g. contoso.onmicrosoft.com or a
    custom verified domain), written to
    HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\CDJ\AAD\TenantName
    REQUIRED when -WriteTargetCdjKeys is $true (the default).

.PARAMETER WriteTargetCdjKeys
    Write the TARGET CDJ registry keys after the leave so the device joins the
    destination tenant. Default: $true

.PARAMETER RemoveCerts
    Remove leftover MS-Organization-Access / MS-Organization-P2P-Access certs
    from LocalMachine\My that 'dsregcmd /leave' can leave behind. Default: $true

.PARAMETER LogDir
    Custom folder for the log file. Defaults to %TEMP%.

.PARAMETER WhatIf
    Standard dry-run. Reports the device state and intended actions, changes nothing.

.EXAMPLE
    .\Invoke-HybridUnjoin.ps1 -TargetTenantName "contoso.onmicrosoft.com" -Verbose
    Guarded unjoin from source, then write target CDJ keys, with console output.

.EXAMPLE
    .\Invoke-HybridUnjoin.ps1 -TargetTenantName "contoso.onmicrosoft.com" -WhatIf
    Dry run - show whether the guard matches and what would happen.

.NOTES
    MUST be run elevated (Administrator). Not designed to run as SYSTEM.

    Separate from OLicenseCleanup.ps1 by design: unjoining a device from its
    tenant is destructive, machine-level, and migration-specific; it must never
    run from a routine login-triggered cleanup task.

    This script does NOT touch the on-prem AD domain join (DomainJoined stays
    YES). It only affects the Azure AD / Entra device registration.

    Verify afterwards with:  dsregcmd /status
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $SourceTenantId     = "1676489c-5c72-46b7-ba63-9ab90c4aad44",
    [string] $TargetTenantId     = "f2ee1ec7-fe58-4178-b8a8-52cc9c5cb34a",
    [string] $TargetTenantName   = "ptcl4.onmicrosoft.com",
    [bool]   $WriteTargetCdjKeys = $true,
    [bool]   $RemoveCerts        = $true,
    [string] $LogDir             = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

#region -- Constants ------------------------------------------------------------
$SCRIPT_VERSION   = "2.0"
$ORG_CERT_ISSUERS = @("MS-Organization-Access", "MS-Organization-P2P-Access")
$CDJ_KEY          = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CDJ\AAD"
#endregion

#region -- Script-scope state ---------------------------------------------------
$Script:LogStream   = $null
$Script:LogFilePath = $null
$Script:TimeStamp   = ""
#endregion

###############################################################################
#region -- Logging helpers (matches OLicenseCleanup.ps1 style) ------------------
###############################################################################

function Write-LogRaw {
    param([string]$Line)
    if ($Script:LogStream) { $Script:LogStream.WriteLine($Line) }
}

function Write-LogHeader {
    param([string]$Message)
    $underline = "=" * $Message.Length
    Write-LogRaw ""
    Write-LogRaw "$Message`r`n$underline"
    Write-Verbose ""
    Write-Verbose $Message
    Write-Verbose $underline
}

function Write-LogSubHeader {
    param([string]$Message)
    $underline = "-" * $Message.Length
    Write-LogRaw ""
    Write-LogRaw "$Message`r`n$underline"
    Write-Verbose ""
    Write-Verbose $Message
    Write-Verbose $underline
}

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
#endregion

###############################################################################
#region -- Setup ----------------------------------------------------------------
###############################################################################

function Test-IsElevated {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $pr = New-Object System.Security.Principal.WindowsPrincipal($id)
    return $pr.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

function New-Log {
    if ($LogDir -eq "") { $LogDir = $env:TEMP }
    $ts = Get-Date -Format "yyyyMMddHHmmss"
    $Script:TimeStamp = $ts
    $name = "$LogDir\$($env:COMPUTERNAME)_${ts}_HybridUnjoin.txt"
    try {
        $Script:LogStream   = [System.IO.StreamWriter]::new($name, $false, [System.Text.Encoding]::Unicode)
        $Script:LogFilePath = $name
    }
    catch {
        $name = "$($env:TEMP)\$($env:COMPUTERNAME)_${ts}_HybridUnjoin.txt"
        $Script:LogStream   = [System.IO.StreamWriter]::new($name, $false, [System.Text.Encoding]::Unicode)
        $Script:LogFilePath = $name
    }

    Write-LogHeader "Hybrid Azure AD Unjoin Utility - v$SCRIPT_VERSION"
    Write-Log "Computer:        $($env:COMPUTERNAME)"
    Write-Log "Run time:        $(Get-Date)"
    Write-Log "Source TenantId: $SourceTenantId"
    Write-Log "Target TenantId: $TargetTenantId"
    Write-Log "Log file:        $($Script:LogFilePath)"
}
#endregion

###############################################################################
#region -- Device-state detection -----------------------------------------------
###############################################################################

#-------------------------------------------------------------------------------
#   Get-DsRegState
#
#   Runs 'dsregcmd /status' and parses Key : Value lines into a hashtable.
#
#   IMPORTANT: dsregcmd /status repeats some keys (notably TenantId and
#   TenantName) in more than one section - e.g. the "Device State" block AND
#   the "SSO State" / User State block, which reflect the logged-on user's PRT
#   rather than the device. A naive last-write-wins parse can therefore return
#   the USER tenant, not the DEVICE tenant. Because the guard must decide based
#   on the DEVICE registration, we capture the value seen in the Device State
#   section into a distinct key, 'DeviceState_TenantId', and use that for the
#   guard. The flat keys remain available for the readable summary.
#-------------------------------------------------------------------------------
function Get-DsRegState {
    $state = @{}
    try {
        $raw = & dsregcmd.exe /status 2>$null
    }
    catch {
        Write-Log "ERROR: could not run dsregcmd.exe: $_"
        return $state
    }

    $section = ""
    foreach ($line in $raw) {
        # Section headers appear as a line of "| Device State |" between rule lines.
        if ($line -match '^\s*\|\s*(.+?)\s*\|\s*$') {
            $section = $matches[1].Trim()
            continue
        }
        if ($line -match '^\s*([A-Za-z0-9_]+)\s*:\s*(.+?)\s*$') {
            $key = $matches[1]
            $val = $matches[2].Trim()
            $state[$key] = $val   # flat (last-write-wins) for the summary
            if ($section -eq "Device State") {
                $state["DeviceState_$key"] = $val   # section-scoped, authoritative
            }
        }
    }
    return $state
}

function Write-DeviceStateSummary {
    param([hashtable]$State)

    $fields = @("AzureAdJoined","EnterpriseJoined","DomainJoined","WorkplaceJoined","DeviceId","TenantName","TenantId")
    Write-LogSubHeader "Device join state (dsregcmd /status)"
    foreach ($f in $fields) {
        if ($State.ContainsKey($f)) {
            Write-Log ("{0,-18}: {1}" -f $f, $State[$f])
        }
    }
}
#endregion

###############################################################################
#region -- Actions --------------------------------------------------------------
###############################################################################

#-------------------------------------------------------------------------------
#   Invoke-DsRegLeave
#-------------------------------------------------------------------------------
function Invoke-DsRegLeave {
    if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, "dsregcmd /leave (unjoin from source tenant)")) {
        Write-Log "Running dsregcmd /leave ..."
        try {
            $dsArgs = if ($VerbosePreference -eq "Continue") { @("/debug","/leave") } else { @("/leave") }
            $output = & dsregcmd.exe @dsArgs 2>&1
            foreach ($l in $output) { Write-Log "   dsreg: $l" }
            Write-Log "dsregcmd /leave completed (exit code $LASTEXITCODE)."
        }
        catch {
            Write-Log "ERROR running dsregcmd /leave: $_"
        }
    }
}

#-------------------------------------------------------------------------------
#   Clear-WamAccounts
#
#   Flushes the WAM account store via dsregcmd /cleanupaccounts when available.
#   Targets the exact layer behind the "another account already signed in" error.
#-------------------------------------------------------------------------------
function Clear-WamAccounts {
    Write-LogSubHeader "Flushing WAM account store (dsregcmd /cleanupaccounts)"
    if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, "dsregcmd /cleanupaccounts")) {
        try {
            $help = (& dsregcmd.exe /? 2>&1) -join "`n"
            if ($help -match "cleanupaccounts") {
                $out = & dsregcmd.exe /cleanupaccounts 2>&1
                foreach ($l in $out) { Write-Log "   dsreg: $l" }
                Write-Log "dsregcmd /cleanupaccounts completed (exit code $LASTEXITCODE)."
            }
            else {
                Write-Log "dsregcmd on this build has no /cleanupaccounts switch - skipping."
            }
        }
        catch {
            Write-Log "WARN: dsregcmd /cleanupaccounts failed: $_"
        }
    }
}

#-------------------------------------------------------------------------------
#   Remove-OrgAccessCerts
#-------------------------------------------------------------------------------
function Remove-OrgAccessCerts {
    Write-LogSubHeader "Removing leftover MS-Organization-* certificates"
    $found = $false
    try {
        $certs = Get-ChildItem "Cert:\LocalMachine\My" -ErrorAction Stop
    }
    catch {
        Write-Log "WARN: could not read LocalMachine\My cert store: $_"
        return
    }

    foreach ($cert in $certs) {
        foreach ($issuer in $ORG_CERT_ISSUERS) {
            if ($cert.Issuer -match [regex]::Escape($issuer)) {
                $found = $true
                $desc = "$($cert.Thumbprint) (issuer: $issuer)"
                if ($PSCmdlet.ShouldProcess($desc, "Remove certificate")) {
                    try {
                        Remove-Item -Path $cert.PSPath -Force -ErrorAction Stop
                        Write-Log "Removed certificate $desc"
                    }
                    catch {
                        Write-Log "WARN: could not remove cert $desc : $_"
                    }
                }
            }
        }
    }
    if (-not $found) { Write-Log "No MS-Organization-* certificates found." }
}

#-------------------------------------------------------------------------------
#   Write-TargetCdjKeys
#
#   Writes the TARGET tenant CDJ keys so Automatic-Device-Join targets the
#   destination tenant (overrides SCP discovery).
#-------------------------------------------------------------------------------
function Write-TargetCdjKeys {
    Write-LogSubHeader "Writing TARGET tenant CDJ keys"

    if ($TargetTenantName -eq "") {
        Write-Log "ERROR: -TargetTenantName is empty. Cannot write CDJ keys without the"
        Write-Log "       target verified domain (e.g. contoso.onmicrosoft.com). Skipping."
        return
    }

    if ($PSCmdlet.ShouldProcess($CDJ_KEY, "Write TenantId + TenantName")) {
        try {
            if (-not (Test-Path $CDJ_KEY)) {
                New-Item -Path $CDJ_KEY -Force -ErrorAction Stop | Out-Null
                Write-Log "Created key: $CDJ_KEY"
            }
            Set-ItemProperty -Path $CDJ_KEY -Name "TenantId"   -Value $TargetTenantId   -Type String -ErrorAction Stop
            Set-ItemProperty -Path $CDJ_KEY -Name "TenantName" -Value $TargetTenantName -Type String -ErrorAction Stop
            Write-Log "Set TenantId   = $TargetTenantId"
            Write-Log "Set TenantName = $TargetTenantName"
        }
        catch {
            Write-Log "ERROR writing CDJ keys: $_"
        }
    }
}
#endregion

###############################################################################
#region -- MAIN -----------------------------------------------------------------
###############################################################################

try {
    if (-not (Test-IsElevated)) {
        Write-Warning "This script must be run as Administrator. Re-launch from an elevated PowerShell prompt."
        return
    }

    New-Log

    # ---- 1. Inspect current state -----------------------------------------
    $state = Get-DsRegState
    if ($state.Count -eq 0) {
        Write-Log "Could not determine device state - aborting to avoid blind changes."
        return
    }
    Write-DeviceStateSummary -State $state

    # Prefer the Device State section's TenantId; fall back to flat if absent.
    $currentTenantId = ""
    if ($state.ContainsKey("DeviceState_TenantId")) {
        $currentTenantId = $state["DeviceState_TenantId"]
    }
    elseif ($state.ContainsKey("TenantId")) {
        $currentTenantId = $state["TenantId"]
        Write-Log "NOTE: Device State TenantId not found; using flat TenantId for the guard."
    }
    Write-Log "Device tenant for guard evaluation: '$currentTenantId'"

    # ---- 2. GUARD: only proceed if joined to the SOURCE tenant ------------
    if ($currentTenantId -ne $SourceTenantId) {
        Write-Log ""
        Write-Log "GUARD: device TenantId '$currentTenantId' does not match source"
        Write-Log "       '$SourceTenantId'. No unjoin performed."
        if ($currentTenantId -eq $TargetTenantId) {
            Write-Log "       Device already appears to be on the TARGET tenant - nothing to do."
        }
        elseif ($currentTenantId -eq "") {
            Write-Log "       Device is not Azure AD joined to any tenant - nothing to do."
        }
        else {
            Write-Log "       Device is on an unexpected tenant - leaving it untouched for safety."
        }
        Write-Host "`nNo changes made (guard did not match source tenant). Log: $($Script:LogFilePath)" -ForegroundColor Yellow
        return
    }

    Write-Log ""
    Write-Log "GUARD MATCHED: device is joined to the SOURCE tenant. Proceeding with unjoin."

    # ---- 3. Leave the source tenant ---------------------------------------
    Invoke-DsRegLeave

    # ---- 4. Flush WAM accounts --------------------------------------------
    Clear-WamAccounts

    # ---- 5. Remove leftover org certs -------------------------------------
    if ($RemoveCerts) {
        Remove-OrgAccessCerts
    }

    # ---- 6. Write target CDJ keys -----------------------------------------
    if ($WriteTargetCdjKeys) {
        Write-TargetCdjKeys
    }

    # ---- 7. Re-check state -------------------------------------------------
    Start-Sleep -Seconds 2
    $after = Get-DsRegState
    Write-DeviceStateSummary -State $after
    $nowJoined = ($after["AzureAdJoined"] -eq "YES")

    $afterTenantId = if ($after.ContainsKey("DeviceState_TenantId")) { $after["DeviceState_TenantId"] }
                     elseif ($after.ContainsKey("TenantId")) { $after["TenantId"] } else { "" }
    Write-LogHeader "Result"
    if ($nowJoined -and $afterTenantId -eq $SourceTenantId) {
        Write-Log "Still reports source tenant. A reboot is likely required for /leave to"
        Write-Log "fully apply; re-check dsregcmd /status after reboot."
    }
    else {
        Write-Log "Device left the source tenant successfully."
    }

    Write-Log ""
    Write-Log "NEXT STEPS:"
    Write-Log "  1. Reboot the device."
    Write-Log "  2. Run OLicenseCleanup.ps1 to clear the (now unseeded) Office caches."
    Write-Log "  3. The Automatic-Device-Join task will hybrid-join the TARGET tenant"
    Write-Log "     (driven by the CDJ keys) once the device object is synced there."
    Write-Log "  4. Open Office and activate with the new (target) account."

    Write-Host "`nHybrid unjoin finished (source -> target). Log: $($Script:LogFilePath)" -ForegroundColor Green
}
finally {
    if ($Script:LogStream) {
        $Script:LogStream.Flush()
        $Script:LogStream.Close()
    }
}
#endregion
