<#
.SYNOPSIS
Disable Microsoft Entra ID and on-premises AD user accounts from a Workday leavers CSV stored in Azure Blob Storage, and produce a report.

.DESCRIPTION
The script downloads the specified CSV from Azure Storage using Managed Identity authentication, matches users by Employee ID, Primary Email, or Legal Name,
and optionally disables matched enabled users in both Microsoft Entra ID and on-premises AD (controlled by the -DisableAccounts switch). It writes an 
augmented CSV and a log file to the local temp directory and sends them via email using SendGrid. The temporary files are not persisted to Azure Storage.

.PARAMETER StorageAccountName
The Azure Storage account name hosting the CSV.

.PARAMETER ContainerName
The blob container name hosting the CSV.

.PARAMETER BlobName
The blob name of the CSV to process (e.g. "Leavers_Report_-_Active_Directory_-_IT.csv").

.PARAMETER OnPremDomainController
FQDN of the on-premises domain controller to target for AD queries and account disable operations.

.PARAMETER OnPremAdUsernameSecretName
Key Vault secret name that contains the on-prem AD username (e.g., UNIPHAR\\AutomationSvc).

.PARAMETER OnPremAdPassSecretName
Key Vault secret name that contains the on-prem AD password corresponding to the username.

.PARAMETER AdKeyVaultName
Azure Key Vault name that stores on-premises Active Directory credentials (username and password secrets). This Key Vault must contain the secrets specified by OnPremAdUsernameSecretName and OnPremAdPassSecretName parameters.

.PARAMETER SecretsKeyVaultName
Azure Key Vault name that stores API keys and external service credentials (SendGrid API key). This Key Vault must contain the secret specified by SendGridApiKeySecretName parameter.

.PARAMETER SendGridApiKeySecretName
The secret name in Key Vault that contains the SendGrid API key (value should be the raw API key).

.PARAMETER SendGridSenderEmailAddress
The sender email address for SendGrid.

.PARAMETER SendGridRecipientEmailAddresses
One or more recipient email addresses for the report.

.PARAMETER SendGridApiEndpoint
SendGrid API endpoint. Default: https://api.sendgrid.com/v3/mail/send

.PARAMETER DisableAccounts
Switch parameter that controls whether account disabling is performed. If false (default), no accounts will be disabled in either Entra ID or on-premises AD. If true, matched accounts will be disabled in both Entra ID and on-premises AD. This parameter acts as a master switch for all disabling operations.

The email subject is fixed: "disabling leavers report".

.EXAMPLE
pwsh -File .\runbooks\Disable-Leavers.ps1 -StorageAccountName unipharsftp -ContainerName workday -BlobName 'Leavers_Report_-_Active_Directory_-_IT.csv' -OnPremDomainController unidc10.uniphar.local -OnPremAdUsernameSecretName <usernameSecret> -OnPremAdPassSecretName <passwordSecret> -AdKeyVaultName <adKeyVaultName> -SecretsKeyVaultName <secretsKeyVaultName> -SendGridApiKeySecretName <sendGridSecret> -SendGridSenderEmailAddress <senderEmail> -SendGridRecipientEmailAddresses <recipientEmails> -DisableAccounts

.REQUIREMENTS
- Microsoft Graph PowerShell with User.ReadWrite.All permissions
- Az.Accounts, Az.Storage, Az.KeyVault modules with Managed Identity access to Storage and Key Vault
- RSAT ActiveDirectory module and Hybrid Runbook Worker network connectivity to on-premises DCs
- SendGrid API key for email notifications
#>

[CmdletBinding(PositionalBinding = $false)]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Enter the Azure Storage account name (e.g., 'unipharsftp')")]
    [string]$StorageAccountName, 
    [Parameter(Mandatory = $true, HelpMessage = "Enter the blob container name (e.g., 'workday')")]
    [string]$ContainerName,
    [Parameter(Mandatory = $true, HelpMessage = "Enter the CSV blob name (e.g., 'Leavers_Report_-_Active_Directory_-_IT.csv')")]
    [string]$BlobName,
    [Parameter(Mandatory = $true, HelpMessage = "Enter the FQDN of the on-premises domain controller (e.g., 'dc117.uni.local')")]
    [string]$OnPremDomainController,
    [Parameter(Mandatory = $true, HelpMessage = "Enter the Key Vault secret name containing the on-prem AD username (e.g., 'OnPremADUser')")]
    [string]$OnPremAdUsernameSecretName,
    [Parameter(Mandatory = $true, HelpMessage = "Enter the Key Vault secret name containing the on-prem AD password (e.g., 'OnPremADPassword')")]
    [string]$OnPremAdPassSecretName,
    [Parameter(Mandatory = $true, HelpMessage = "Enter the Key Vault name for AD credentials (e.g., 'UniPharADKeyVault') - stores on-premises AD username and password secrets")]
    [string]$AdKeyVaultName,
    [Parameter(Mandatory = $true, HelpMessage = "Enter the Key Vault name for API credentials (e.g., 'UniPharSecretsKeyVault') - stores SendGrid API key and other external service credentials")]
    [string]$SecretsKeyVaultName,
    [Parameter(Mandatory = $true, HelpMessage = "Enter the Key Vault secret name containing the SendGrid API key (e.g., 'SendGridApiKey')")]
    [string]$SendGridApiKeySecretName,
    [Parameter(Mandatory = $true, HelpMessage = "Enter the sender email address for SendGrid (e.g., 'noreply@uniphar.ie')")]
    [string]$SendGridSenderEmailAddress,
    [Parameter(Mandatory = $true, HelpMessage = "Enter recipient email addresses separated by commas (e.g., 'it@uniphar.com,admin@uniphar.com')")]
    [string[]]$SendGridRecipientEmailAddresses,
    [Parameter(Mandatory = $false, HelpMessage = "Enter SendGrid API endpoint URL (default: https://api.sendgrid.com/v3/mail/send)")]
    [string]$SendGridApiEndpoint = 'https://api.sendgrid.com/v3/mail/send',
    [Parameter(Mandatory = $false, HelpMessage = "Check this box to actually disable user accounts. Leave unchecked for reports only (safe mode)")]
    [switch]$DisableAccounts = $false
)

# Disable-Leavers script - Process Workday leavers and disable accounts
# Resolve system temp directory and ensure it exists
$LocalTempDir = [System.IO.Path]::GetTempPath()
if (-not (Test-Path $LocalTempDir)) { New-Item -ItemType Directory -Path $LocalTempDir -Force | Out-Null }

# Timestamp once per run so CSVs are unique and not overwritten
$ReportFilenamePattern = "Leavers_Report_-_Active_Directory_-_IT_withUPN"
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$OutputPath = Join-Path $LocalTempDir "${ReportFilenamePattern}_${timestamp}.csv"

# Initialize Azure and Graph contexts (Managed Identity when available)
Initialize-AzureContext
Initialize-GraphContext

# Build Azure Storage context using managed identity / connected account
try {
    $storageContext = New-AzStorageContext -StorageAccountName $StorageAccountName -UseConnectedAccount -ErrorAction Stop
}
catch {
    throw "Failed to create storage context: $($_.Exception.Message)"
}
$InputPath = Join-Path $LocalTempDir (Split-Path -Leaf $BlobName)
Write-Host "Downloading input from Azure Storage: $StorageAccountName/$ContainerName/$BlobName -> $InputPath" -ForegroundColor Cyan
Get-AzStorageBlobContent -Container $ContainerName -Blob $BlobName -Destination $InputPath -Context $storageContext -Force | Out-Null

# Graph context is already initialized above via Initialize-GraphContext

Write-Host "Loading input CSV: $InputPath" -ForegroundColor Cyan
$rows = Import-Csv -Path $InputPath
if (-not $rows) {
    Write-Warning "No rows loaded from CSV. Exiting."; return
}

Write-Host "Retrieving users from Entra ID (this may take time in large tenants)..." -ForegroundColor Cyan
# Single directory pull (avoid N x API calls). Adjust properties as needed.
$allUsers = Get-MgUser -All -Property id, employeeId, mail, displayName, userPrincipalName, accountEnabled -ConsistencyLevel eventual |
Select-Object id, userPrincipalName, employeeId, mail, displayName, accountEnabled

# Build fast lookup hash tables (case-insensitive keys)
$byEmployeeId = @{}
$byMail = @{}
$byDisplay = @{}
foreach ($u in $allUsers) {
    if ($u.employeeId -and -not $byEmployeeId.ContainsKey($u.employeeId)) { $byEmployeeId[$u.employeeId] = $u }
    if ($u.mail -and -not $byMail.ContainsKey($u.mail.ToLower())) { $byMail[$u.mail.ToLower()] = $u }
    if ($u.displayName) {
        $dnKey = $u.displayName.ToLower()
        if (-not $byDisplay.ContainsKey($dnKey)) { $byDisplay[$dnKey] = $u }
    }
}

$processed = 0
$total = $rows.Count
foreach ($row in $rows) {
    $processed++
    $employeeId = $row.'Employee ID'
    $email = $row.'Email - Primary Work'
    # Current file uses 'Legal Name' for display name source
    $displayName = $row.'Legal Name'

    # Ensure / (re)create output columns (idempotent with -Force)
    Add-Member -InputObject $row -NotePropertyName UPN             -NotePropertyValue $null -Force
    Add-Member -InputObject $row -NotePropertyName 'UPN-mail'      -NotePropertyValue $null -Force
    Add-Member -InputObject $row -NotePropertyName DisplayName_UPN -NotePropertyValue $null -Force
    Add-Member -InputObject $row -NotePropertyName AccountEnabled  -NotePropertyValue $null -Force
    Add-Member -InputObject $row -NotePropertyName MatchSource     -NotePropertyValue $null -Force
    Add-Member -InputObject $row -NotePropertyName OnPremDisabledActionResult -NotePropertyValue $null -Force

    $user = $null

    # 1. Match by employeeId (exact)
    if ($employeeId -and $byEmployeeId.ContainsKey($employeeId)) {
        $user = $byEmployeeId[$employeeId]
        $row.UPN = $user.userPrincipalName
        $row.AccountEnabled = $user.accountEnabled
        $row.MatchSource = 'EmployeeId'
    }
    # 2. Match by mail (case-insensitive exact)
    elseif ($email) {
        $lower = $email.ToLower()
        if ($byMail.ContainsKey($lower)) {
            $user = $byMail[$lower]
            $row.'UPN-mail' = $user.userPrincipalName
            $row.AccountEnabled = $user.accountEnabled
            $row.MatchSource = 'Mail'
        }
    }
    # 3. Match by display name (case-insensitive exact)
    if (-not $user -and $displayName) {
        $dnKey = $displayName.ToLower()
        if ($byDisplay.ContainsKey($dnKey)) {
            $user = $byDisplay[$dnKey]
            $row.DisplayName_UPN = $user.userPrincipalName
            $row.AccountEnabled = $user.accountEnabled
            $row.MatchSource = 'DisplayName'
        }
    }

    if (-not $user) {
        # Leave columns null; could optionally log
        # Write-Verbose "No match for row $processed/$total (EmployeeId=$employeeId, Email=$email, DisplayName=$displayName)"
    }

    if (($processed % 200) -eq 0) { Write-Host "Processed $processed / $total" -ForegroundColor DarkGray }
}

# After matching loop, disable matched enabled accounts if DisableAccounts switch is enabled
$rows | Add-Member -NotePropertyName DisabledActionResult -NotePropertyValue $null -Force

if ($DisableAccounts) {
    Write-Host "Disabling matched enabled cloud accounts..." -ForegroundColor Yellow
    $matchedEnabled = $rows | Where-Object { $_.MatchSource -and $_.AccountEnabled -eq $true }

    # Prepare on-prem AD context early if requested
    if (-not (Initialize-OnPremAD -Server $OnPremDomainController)) {
        Write-Warning "On-prem AD initialization failed. Only Entra ID accounts will be disabled."
        $onPremAvailable = $false
    } else {
        $onPremAvailable = $true
    }

    foreach ($r in $matchedEnabled) {
        $upnToDisable = $r.UPN
        if (-not $upnToDisable) { $upnToDisable = $r.'UPN-mail' }
        if (-not $upnToDisable) { $upnToDisable = $r.DisplayName_UPN }
        if (-not $upnToDisable) { continue }
        try {
            Update-MgUser -UserId $upnToDisable -AccountEnabled:$false
            $r.DisabledActionResult = 'Disabled'
        }
        catch {
            $r.DisabledActionResult = "Error: $($_.Exception.Message)"
        }
    }

    # On-prem disable ALL matched accounts (even those already disabled in Entra)
    if ($onPremAvailable) {
        Write-Host "Disabling matched accounts on-prem (all matched, regardless of cloud state)..." -ForegroundColor Yellow
        $matchedAll = $rows | Where-Object { $_.MatchSource }
        foreach ($r in $matchedAll) {
            $upnToDisable = $r.UPN
            if (-not $upnToDisable) { $upnToDisable = $r.'UPN-mail' }
            if (-not $upnToDisable) { $upnToDisable = $r.DisplayName_UPN }
            if (-not $upnToDisable) { continue }
            try {
                $adUserParams = @{ Identity = $upnToDisable; Server = $OnPremDomainController; Properties = 'Enabled'; ErrorAction = 'Stop' }
                if ($global:AdCredential) { $adUserParams['Credential'] = $global:AdCredential }
                $adUser = $null
                $adUser = Get-ADUser @adUserParams
                if ($adUser.Enabled) {
                    $disParams = @{ Identity = $adUser.DistinguishedName; Server = $OnPremDomainController; ErrorAction = 'Stop' }
                    if ($global:AdCredential) { $disParams['Credential'] = $global:AdCredential }
                    Disable-ADAccount @disParams
                    $r.OnPremDisabledActionResult = 'Disabled'
                }
                else {
                    # Attempting to disable again but it's already disabled
                    if (-not $r.OnPremDisabledActionResult) { $r.OnPremDisabledActionResult = 'AlreadyDisabled' }
                }
            }
            catch {
                $r.OnPremDisabledActionResult = "Error: $($_.Exception.Message)"
            }
        }
    }
}
else {
    Write-Host "DisableAccounts is false - no accounts will be disabled" -ForegroundColor Yellow
}

# Reporting setup (reuse same $timestamp as output file)
$LogPath = (Join-Path $LocalTempDir "Disable-Leavers_Report_$timestamp.log")
# Ensure output and log directories exist
$__pathsToEnsure = @((Split-Path -Parent $OutputPath), (Split-Path -Parent $LogPath))
foreach ($__d in $__pathsToEnsure) { if ($__d -and -not (Test-Path $__d)) { New-Item -ItemType Directory -Path $__d -Force | Out-Null } }
"Disable-Leavers run started: $(Get-Date)" | Out-File -FilePath $LogPath -Encoding UTF8
"Input CSV: $InputPath" | Out-File -FilePath $LogPath -Append
"Output CSV (will be written): $OutputPath" | Out-File -FilePath $LogPath -Append
"DisableAccounts parameter: $DisableAccounts" | Out-File -FilePath $LogPath -Append

Write-Host "Writing output CSV: $OutputPath" -ForegroundColor Cyan
$rows | Export-Csv -Path $OutputPath -NoTypeInformation

# Build list of files to send (kept in temp directory)
$filesToSend = @($InputPath, $OutputPath, $LogPath) | Where-Object { $_ -and (Test-Path $_) }

# Build report summary
$matched = $rows | Where-Object { $_.MatchSource }
$matchedCount = $matched.Count
$unknownCount = $rows.Count - $matchedCount
$bySource = $matched | Group-Object MatchSource | Select-Object Name, Count
$disabledSuccess = @($rows | Where-Object { $_.DisabledActionResult -eq 'Disabled' })
$disabledErrors = @($rows | Where-Object { $_.DisabledActionResult -like 'Error:*' })
$onPremDisabled = @($rows | Where-Object { $_.OnPremDisabledActionResult -eq 'Disabled' })
$onPremDisableErr = @($rows | Where-Object { $_.OnPremDisabledActionResult -like 'Error:*' })
$disabledCount = $disabledSuccess.Count
$disableErrorCount = $disabledErrors.Count

"" | Out-File -FilePath $LogPath -Append
"Summary:" | Out-File -FilePath $LogPath -Append
"Total rows:        $($rows.Count)" | Out-File -FilePath $LogPath -Append
"Matched rows:      $matchedCount" | Out-File -FilePath $LogPath -Append
foreach ($g in $bySource) { "  Matched by $($g.Name): $($g.Count)" | Out-File -FilePath $LogPath -Append }
"Unknown rows:      $unknownCount" | Out-File -FilePath $LogPath -Append
if ($DisableAccounts) {
    "Accounts disabled: $disabledCount" | Out-File -FilePath $LogPath -Append
    "On-prem accounts disabled: $($onPremDisabled.Count)" | Out-File -FilePath $LogPath -Append
    if ($disableErrorCount -gt 0) { "Disable errors:   $disableErrorCount" | Out-File -FilePath $LogPath -Append }
    if ($onPremDisableErr.Count -gt 0) { "On-prem disable errors: $($onPremDisableErr.Count)" | Out-File -FilePath $LogPath -Append }
}
else {
    "DisableAccounts is false - no accounts were disabled" | Out-File -FilePath $LogPath -Append
}

if ($DisableAccounts -and $disableErrorCount -gt 0) {
    "" | Out-File -FilePath $LogPath -Append
    "Disable error details:" | Out-File -FilePath $LogPath -Append
    foreach ($e in $disabledErrors) {
        $u = $e.UPN; if (-not $u) { $u = $e.'UPN-mail' }; if (-not $u) { $u = $e.DisplayName_UPN }
        "  Cloud: $u => $($e.DisabledActionResult)" | Out-File -FilePath $LogPath -Append
    }
}
if ($DisableAccounts -and $onPremDisableErr.Count -gt 0) {
    "" | Out-File -FilePath $LogPath -Append
    "On-prem disable error details:" | Out-File -FilePath $LogPath -Append
    foreach ($e in $onPremDisableErr) {
        $u = $e.UPN; if (-not $u) { $u = $e.'UPN-mail' }; if (-not $u) { $u = $e.DisplayName_UPN }
        "  OnPrem: $u => $($e.OnPremDisabledActionResult)" | Out-File -FilePath $LogPath -Append
    }
}

"Run completed: $(Get-Date)" | Out-File -FilePath $LogPath -Append
Write-Host "Report log written: $LogPath" -ForegroundColor Green


#all the functions:

# Fetch SendGrid API key from Key Vault and send email with attachments
function Get-SecretFromKeyVault {
    param([string]$VaultName, [string]$SecretName)
    try {
        return (Get-AzKeyVaultSecret -VaultName $VaultName -Name $SecretName -AsPlainText -ErrorAction Stop)
    }
    catch {
        throw "Unable to retrieve secret '$SecretName' from Key Vault '$VaultName': $($_.Exception.Message)"
    }
}

function Get-MimeTypeForFile {
    param([string]$Path)
    $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    switch ($ext) {
        '.csv' { 'text/csv'; break }
        '.log' { 'text/plain'; break }
        '.txt' { 'text/plain'; break }
        default { 'application/octet-stream' }
    }
}

function Send-ReportViaSendGrid {
    param(
        [string]$ApiKey,
        [string]$FromEmail,
        [string[]]$Recipients,
        [string]$Subject,
        [string]$Endpoint,
        [string[]]$AttachmentPaths
    )
    $toArray = @(); foreach ($r in $Recipients) { if (-not [string]::IsNullOrWhiteSpace($r)) { $toArray += @{ email = $r.Trim() } } }
    if ($toArray.Count -eq 0) { throw 'No valid recipient email addresses specified.' }
    $attachments = @()
    foreach ($p in $AttachmentPaths) {
        try {
            $bytes = [System.IO.File]::ReadAllBytes($p)
            $b64 = [System.Convert]::ToBase64String($bytes)
            $attachments += @{ content = $b64; filename = (Split-Path -Leaf $p); type = (Get-MimeTypeForFile -Path $p); disposition = 'attachment' }
        }
        catch {
            Write-Warning "Failed to attach file '$p': $($_.Exception.Message)"
        }
    }
    $bodyObj = @{ 
        personalizations = @(@{ to = $toArray; subject = $Subject })
        from             = @{ email = $FromEmail }
        content          = @(@{ type = 'text/plain'; value = "Leavers disable run completed at $(Get-Date). See attached output CSV and log." })
        attachments      = $attachments
    }
    $headers = @{ Authorization = "Bearer $ApiKey" }
    $json = $bodyObj | ConvertTo-Json -Depth 10
    try {
        Invoke-RestMethod -Method Post -Uri $Endpoint -Headers $headers -Body $json -ContentType 'application/json' -ErrorAction Stop | Out-Null
        Write-Host 'SendGrid report email sent.' -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to send SendGrid email: $($_.Exception.Message)"
    }
}

try {
    # Use the specified Secrets Key Vault directly
    $apiKey = Get-SecretFromKeyVault -VaultName $SecretsKeyVaultName -SecretName $SendGridApiKeySecretName
    $fixedSubject = 'disabling leavers report'
    Send-ReportViaSendGrid -ApiKey $apiKey -FromEmail $SendGridSenderEmailAddress -Recipients $SendGridRecipientEmailAddresses -Subject $fixedSubject -Endpoint $SendGridApiEndpoint -AttachmentPaths $filesToSend
}
catch {
    Write-Warning $_
}

function Initialize-AzureContext {
    try {
        Import-Module Az.Accounts -ErrorAction Stop
        Import-Module Az.Storage -ErrorAction Stop
        Import-Module Az.KeyVault -ErrorAction Stop
    }
    catch {
        throw 'Required modules Az.Accounts and Az.Storage are not installed in this environment.'
    }
    $ctx = $null
    try { $ctx = Get-AzContext -ErrorAction Stop } catch { $ctx = $null }
    if (-not $ctx) {
        try {
            # Prefer Managed Identity (Azure Automation / managed hosts)
            Connect-AzAccount -Identity -ErrorAction Stop | Out-Null
        }
        catch {
            # Fallback to interactive (useful for local runs)
            Connect-AzAccount -ErrorAction Stop | Out-Null
        }
    }
}

function Initialize-GraphContext {
    try {
        Import-Module Microsoft.Graph.Users -ErrorAction Stop
    }
    catch {
        throw 'Required module Microsoft.Graph.Users is not installed in this environment.'
    }
    try {
        Connect-MgGraph -Identity -NoWelcome -ErrorAction Stop | Out-Null
    }
    catch {
        throw "Failed to connect to Microsoft Graph with Managed Identity. Ensure the RunAs account has required permissions (User.ReadWrite.All) and is enabled. Error: $($_.Exception.Message)"
    }
}
# Centralized AD initialization
function Initialize-OnPremAD {
    param(
        [string]$Server
    )
    if (-not $Server) { $Server = $OnPremDomainController }
    if ($script:OnPremADReady) { return $true }
    try {
        if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) { throw 'ActiveDirectory module not found (RSAT not installed?)' }
        if ($PSVersionTable.PSEdition -eq 'Core') {
            # Use WindowsCompatibility layer to load AD module in PS7 without switching shells
            Import-Module ActiveDirectory -UseWindowsPowerShell -ErrorAction Stop | Out-Null
        }
        else {
            Import-Module ActiveDirectory -ErrorAction Stop | Out-Null
        }
        # Retrieve credentials from Key Vault
        if (-not $global:AdCredential) {
            try {
                # Use the specified AD Key Vault directly
                $adUser = Get-AzKeyVaultSecret -VaultName $AdKeyVaultName -Name $OnPremAdUsernameSecretName -AsPlainText -ErrorAction Stop
                $adPassPlain = Get-AzKeyVaultSecret -VaultName $AdKeyVaultName -Name $OnPremAdPassSecretName -AsPlainText -ErrorAction Stop
                if ([string]::IsNullOrWhiteSpace($adUser) -or [string]::IsNullOrWhiteSpace($adPassPlain)) { throw 'Empty AD username or password from Key Vault' }
                $adPass = ConvertTo-SecureString $adPassPlain -AsPlainText -Force
                $global:AdCredential = New-Object System.Management.Automation.PSCredential ($adUser, $adPass)
            }
            catch {
                throw "Failed to retrieve on-prem AD credentials from Key Vault: $($_.Exception.Message). On-prem actions cannot proceed."
            }
        }
        # Try simple query to validate connection and credentials
        $dcParams = @{ Server = $Server; ErrorAction = 'Stop' }
        if ($global:AdCredential) { $dcParams['Credential'] = $global:AdCredential }
        Get-ADDomainController @dcParams | Out-Null

        $script:OnPremADReady = $true
        Write-Host 'On-prem AD connectivity OK.' -ForegroundColor Green
        return $true
    }
    catch {
        Write-Warning "On-prem AD not available: $($_.Exception.Message)"
        $script:OnPremADReady = $false
        return $false
    }
}
