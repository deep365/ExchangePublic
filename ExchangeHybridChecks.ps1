#### Exchange Hybrid Checks

$ScriptPath = Split-Path -parent $myInvocation.myCommand.definition # Get current path 
#$ScriptPath = "C:\Code"
$ScriptName = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
#$ScriptName = "ExchangeHybridChecks"
$WorkingDir = $ScriptPath
$detailDate="{0:yyyy_MM_dd-HH-mm-ss}" -f (Get-Date)
$logfile="$WorkingDir\$($ScriptName)_LOG_{0}.log" -f $detailDate
$OutFile = "$ScriptPath\$($ScriptName)_{0}.csv" -f $detailDate
$HostnameExchange = "exchange.contoso.com" #Hostname Exchange Onpremises

function logit($text,$color="yellow") {
    $text = "{0:yyyy-MM-dd HH-mm-ss.fff}`t{1}" -f (Get-Date),$text
    write-host -foregroundcolor $color $text
    $text  | Out-File $logfile -append
}  
function endit() {
    # Purpose: Disconnect from all PS modules and exit
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Remove-PSSession $Session
    #exit
}

If ($Error) {$Error.Clear()}
$startTime = Get-Date
$Watch = [System.Diagnostics.Stopwatch]::StartNew()


logit "######################### Avvio script $ScriptName ##########################"
logit "Procedura per il controllo Ambiente Exchange hybrid"
logit "Computer corrente : $($env:computername).$($env:userdnsdomain) - User $($env:UserDomain)\$($env:UserName)"

#ENABLE TLS 1.2 for this Powershell session
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Connessione ad Exchange Onprem
logit "Connessione ad Exchange On-Premise $HostnameExchange"
$Cred = Get-Credential -Message "Immettere le credenziali amministrative di Exchange On-Premises"
#$Pso = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck #-ProxyAccessType NoProxyServer
$error.Clear()
$Session= New-PSSession -Configuration Microsoft.Exchange -ConnectionUri http://$HostnameExchange/PowerShell/ -Credential $Cred -Authentication Kerberos #-SessionOption $Pso
Import-PSSession $Session -DisableNameChecking -AllowClobber -Prefix "ONPR"
if ($error.Count -ne 0) {
    logit "Unable to connect to Exchange Onprem. Exiting." red
    endit
}
Set-ONPRAdServerSettings -ViewEntireForest $True

# Connettersi ad Exchange Online se non già  connessi
if (!((Get-ConnectionInformation).state -eq 'Connected')){
    logit "Connessione ad Exchange Online"
    $error.Clear()
    Connect-ExchangeOnline -ShowBanner:$False
    if ($error.Count -ne 0) {
        logit "Unable to connect to Exchange Online. Exiting." red
        endit
    }
}


$ONPRRecipients = Get-ONPRRecipient -ResultSize Unlimited
logit "Count of onprem recipients: $($ONPRRecipients.count)"
$ONPRRecs = $ONPRRecipients | select PrimarySmtpAddress, Identity, Name, DisplayName, ExternalEmailAddress, RecipientType, RecipientTypeDetails

$EXORecipients = Get-Recipient -ResultSize Unlimited
logit "Count of EXO recipients: $($EXORecipients.count)"
$EXORecs = $EXORecipients | select PrimarySmtpAddress, Identity, Name, DisplayName, ExternalEmailAddress, RecipientType, RecipientTypeDetails


###### VERSION With IndexOf ######

$Out = New-Object System.Collections.Generic.List[System.Object]
$FoundPrimaries = New-Object System.Collections.Generic.List[System.Object]
#[Array] $Out = @()
#[Array] $FoundPrimaries = @()
$Total = $ONPRRecipients.count
$Count = 0

foreach ($ONPRRecipient in $ONPRRecipients) {
    $rowIndex = $EXORecs.PrimarySmtpAddress.IndexOf($ONPRRecipient.PrimarySmtpAddress)
    if ($rowIndex -ne -1) {
        $Search = $EXORecs[$rowIndex]
        #$FoundPrimaries += $ONPRRecipient.PrimarySmtpAddress
        $FoundPrimaries.Add($ONPRRecipient.PrimarySmtpAddress)
        $Properties = [ordered]@{
            Matched = $True
            Identity = $ONPRRecipient.Identity
            Name = $ONPRRecipient.Name
            DisplayName = $ONPRRecipient.DisplayName
            PrimarySmtpAddress = $ONPRRecipient.PrimarySmtpAddress
            ExternalEmailAddress = $ONPRRecipient.ExternalEmailAddress
            RecipientType = $ONPRRecipient.RecipientType
            RecipientTypeDetails = $ONPRRecipient.RecipientTypeDetails
            EXOIdentity = $Search.Identity
            EXOName = $Search.Name
            EXODisplayName = $Search.DisplayName
            EXOPrimarySmtpAddress = $Search.PrimarySmtpAddress
            EXOExternalEmailAddress = $Search.ExternalEmailAddress
            EXORecipientType = $Search.RecipientType
            EXORecipientTypeDetails = $Search.RecipientTypeDetails
        }
        $item = New-Object PSObject -Property $Properties
        $Out.Add($item)

    } Else {
        $Properties = [ordered]@{
            Matched = $False
            Identity = $ONPRRecipient.Identity
            Name = $ONPRRecipient.Name
            DisplayName = $ONPRRecipient.DisplayName
            PrimarySmtpAddress = $ONPRRecipient.PrimarySmtpAddress
            ExternalEmailAddress = $ONPRRecipient.ExternalEmailAddress
            RecipientType = $ONPRRecipient.RecipientType
            RecipientTypeDetails = $ONPRRecipient.RecipientTypeDetails
            EXOIdentity = ""
            EXOName = ""
            EXODisplayName = ""
            EXOPrimarySmtpAddress = ""
            EXOExternalEmailAddress = ""
            EXORecipientType = ""
            EXORecipientTypeDetails = ""
        }
        $item = New-Object PSObject -Property $Properties
        $Out.Add($item)
    }
    $Count +=1
    if (($Count % 10) -eq 0 ) {
        $Perc = [math]::Round($Count/$Total*100,2)
        Write-Progress -Activity "STEP1 - Matching Onrepm --> EXO.." -Status "$Perc % Complete:" -PercentComplete $Perc
    }
}


#Next we start from EXO recipients that didn't matched previously and add as unmatched
$Total = $EXORecipients.count
$Count = 0
foreach ($EXORecipient in $EXORecipients) {
    $rowIndex = $ONPRRecs.PrimarySmtpAddress.IndexOf($EXORecipient.PrimarySmtpAddress)
    if ($rowIndex -eq -1) {
        $Properties = [ordered]@{
            Matched = $False
            Identity = ""
            Name = ""
            DisplayName = ""
            PrimarySmtpAddress = ""
            ExternalEmailAddress = ""
            RecipientType = ""
            RecipientTypeDetails = ""
            EXOIdentity = $EXORecipient.Identity
            EXOName = $EXORecipient.Name
            EXODisplayName = $EXORecipient.DisplayName
            EXOPrimarySmtpAddress = $EXORecipient.PrimarySmtpAddress
            EXOExternalEmailAddress = $EXORecipient.ExternalEmailAddress
            EXORecipientType = $EXORecipient.RecipientType
            EXORecipientTypeDetails = $EXORecipient.RecipientTypeDetails
        }
        $item = New-Object PSObject -Property $Properties
        $Out.Add($item)
    }
    $Count +=1
    if (($Count % 10) -eq 0 ) {
        $Perc = [math]::Round($Count/$Total*100,2)
        Write-Progress -Activity "STEP2 - Adding unmatched EXO --> Onprem.." -Status "$Perc % Complete:" -PercentComplete $Perc
    }
}

logit "=== Salvo i dati raccolti nel file: $OutFile"
$Out | Export-Csv -NoTypeInformation $Outfile -Delimiter ";"

$endTime     = get-date
$elapsed = [math]::Round($watch.Elapsed.TotalSeconds,0)

"" #blank line
logit "-------------------------------------------------"
logit "Script started at:   $startTime"
logit "Script completed at: $endTime"
logit "Script took $elapsed seconds"
logit "Mailbox totali: $Total"
logit "-------------------------------------------------"
"" #blank line

endit
