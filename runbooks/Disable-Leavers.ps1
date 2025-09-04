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
- Microsoft Graph PowerShell with User.connReadWrite.All permissions
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
    $SendGridRecipientEmailAddresses,
    [Parameter(Mandatory = $false, HelpMessage = "Enter SendGrid API endpoint URL (default: https://api.sendgrid.com/v3/mail/send)")]
    [string]$SendGridApiEndpoint = 'https://api.sendgrid.com/v3/mail/send',
    [Parameter(Mandatory = $false, HelpMessage = "Check this box to actually disable user accounts. Leave unchecked for reports only (safe mode)")]
    $DisableAccounts = $false
)

# Fail fast on non-terminating errors; override per-call if needed
$ErrorActionPreference = 'Stop'

# --- Function definitions moved here (all functions) ---

# Ensure required modules are installed (best-effort, non-interactive)
function Ensure-RequiredModules {
    param(
        [string[]]$ModuleNames = @('Az.Accounts','Az.Storage','Az.KeyVault','Microsoft.Graph')
    )

    Write-Host "Checking and installing required PowerShell modules: $($ModuleNames -join ', ')" -ForegroundColor Cyan

    # Ensure NuGet provider available for older systems
    try {
        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -Force -Scope CurrentUser -ErrorAction SilentlyContinue | Out-Null
        }
    }
    catch {
        Write-Warning "Could not ensure NuGet provider: $($_.Exception.Message)"
    }

    # Ensure PSGallery repository exists
    try {
        $psGallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if (-not $psGallery) {
            Write-Host 'Registering PSGallery repository (untrusted by default).' -ForegroundColor Yellow
            Register-PSRepository -Default -ErrorAction SilentlyContinue | Out-Null
        }
    }
    catch {
        Write-Warning "Could not register PSGallery repository: $($_.Exception.Message)"
    }

    foreach ($m in $ModuleNames) {
        try {
            if (Get-Module -ListAvailable -Name $m -ErrorAction SilentlyContinue) {
                Write-Host "Module '$m' already available." -ForegroundColor DarkGray
                continue
            }

            Write-Host "Installing module '$m' to CurrentUser scope (best-effort)." -ForegroundColor Cyan
            Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop | Out-Null
            Write-Host "Installed module '$m'." -ForegroundColor Green
        }
        catch {
            Write-Warning "Failed to install module '$m': $($_.Exception.Message)"
        }
    }

    # ActiveDirectory (RSAT) is not always available via PSGallery. Try Windows capability as a best-effort on Windows hosts.
    try {
        if (-not (Get-Module -ListAvailable -Name ActiveDirectory -ErrorAction SilentlyContinue)) {
            if ($IsWindows -and (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue)) {
                Write-Host 'ActiveDirectory module not found. Attempting to enable RSAT ActiveDirectory capability (requires admin).' -ForegroundColor Yellow
                try {
                    Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0' -ErrorAction Stop | Out-Null
                    Write-Host 'Requested RSAT ActiveDirectory capability install. Module may be available after reboot/logon.' -ForegroundColor Green
                }
                catch {
                    Write-Warning "Could not install RSAT ActiveDirectory via Add-WindowsCapability: $($_.Exception.Message)"
                }
            }
            else {
                Write-Warning 'ActiveDirectory module not found. On non-Windows or restricted hosts this must be provided by the environment (RSAT or Hybrid Worker image).'
            }
        }
    }
    catch {
        Write-Warning "ActiveDirectory availability check failed: $($_.Exception.Message)"
    }
}

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
        $Recipients,
        [string]$Subject,
        [string]$Endpoint,
        $AttachmentPaths
    )
    Write-Host "SendGrid function called." -ForegroundColor Gray

    # --- Normalize Recipients into an array of PSCustomObject with an 'email' property ---
    $RecipientArray = @()
    if ($null -eq $Recipients) {
        $RecipientArray = @()
    }
    elseif ($Recipients -is [string]) {
        # Accept comma, semicolon or newline separated lists
        $RecipientArray = ($Recipients -split '[,;\r\n]') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    }
    elseif ($Recipients -is [System.Collections.IEnumerable]) {
        $RecipientArray = $Recipients | ForEach-Object { $_.ToString().Trim() } | Where-Object { $_ }
    }
    else {
        $RecipientArray = @($Recipients.ToString().Trim())
    }

    $toarray = @()
    foreach ($r in $RecipientArray) {
        if (-not [string]::IsNullOrWhiteSpace($r)) {
            $toarray += [PSCustomObject]@{ email = $r }
        }
    }
    if ($toarray.Count -eq 0) { throw 'No valid recipient email addresses specified.' }

    # --- Normalize AttachmentPaths into an array of strings ---
    $AttachmentArray = @()
    if ($null -ne $AttachmentPaths) {
        if ($AttachmentPaths -is [string]) { $AttachmentArray = @($AttachmentPaths) }
        elseif ($AttachmentPaths -is [System.Collections.IEnumerable]) { $AttachmentArray = $AttachmentPaths }
        else { $AttachmentArray = @($AttachmentPaths.ToString()) }
    }

    $attachments = @()
    foreach ($p in $AttachmentArray) {
        try {
            if (-not (Test-Path $p)) { Write-Warning "Attachment path not found: $p"; continue }
            $bytes = [System.IO.File]::ReadAllBytes($p)
            $b64 = [System.Convert]::ToBase64String($bytes)
            $attachments += [PSCustomObject]@{
                content     = $b64
                filename    = (Split-Path -Leaf $p)
                type        = (Get-MimeTypeForFile -Path $p)
                disposition = 'attachment'
            }
        }
        catch {
            Write-Warning "Failed to attach file '$p': $($_.Exception.Message)"
        }
    }
    $bodyobj = @{ 
        personalizations = @(@{ to = $toarray; subject = $Subject })
        from             = @{ email = $FromEmail }
        content          = @(@{ type = 'text/plain'; value = "Leavers disable run completed at $(Get-Date). See attached output CSV and log." })
        attachments      = $attachments
    }
    $headers = @{ Authorization = "Bearer $ApiKey" }
    $json = $bodyobj | ConvertTo-Json -Depth 10
    try {
        Invoke-RestMethod -Method Post -Uri $Endpoint -Headers $headers -Body $json -ContentType 'application/json' -ErrorAction Stop | Out-Null
        Write-Host 'SendGrid report email sent.' -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to send SendGrid email: $($_.Exception.Message)"
    }
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
    # Retrieve credentials (lazy) and validate connection
    $dcparams = @{ Server = $Server; ErrorAction = 'Stop' }
    $cred = Get-OnPremAdCredential
    if ($cred) { $dcparams['Credential'] = $cred }
        Get-ADDomainController @dcparams | Out-Null

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

# Lazily retrieve and cache on-prem AD credentials (script scope)
function Get-OnPremAdCredential {
    if ($script:AdCredential) { return $script:AdCredential }
    try {
    $aduser = Get-AzKeyVaultSecret -VaultName $AdKeyVaultName -Name $OnPremAdUsernameSecretName -AsPlainText -ErrorAction Stop
    $adpassplain = Get-AzKeyVaultSecret -VaultName $AdKeyVaultName -Name $OnPremAdPassSecretName -AsPlainText -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($aduser) -or [string]::IsNullOrWhiteSpace($adpassplain)) { throw 'Empty AD username or password from Key Vault' }
    $adpass = ConvertTo-SecureString $adpassplain -AsPlainText -Force
    $script:AdCredential = New-Object System.Management.Automation.PSCredential ($aduser, $adpass)
        return $script:AdCredential
    }
    catch {
        throw "Failed to retrieve on-prem AD credentials from Key Vault: $($_.Exception.Message)."
    }
}

# Simple retry helper with exponential backoff and optional Retry-After header support
function Invoke-WithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ScriptBlock]$Script,
        [int]$MaxAttempts = 5,
        [int]$BaseDelaySeconds = 1
    )
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return & $Script
        }
        catch {
            if ($attempt -ge $MaxAttempts) { throw }
            $ex = $_.Exception
            # Default exponential backoff
            $delay = [math]::Min(60, [int]([math]::Pow(2, $attempt - 1) * $BaseDelaySeconds))
            # If error contains Retry-After, honor it when larger
            $retryAfter = $null
            if ($ex.Response -and $ex.Response.Headers -and $ex.Response.Headers['Retry-After']) {
                [int]::TryParse($ex.Response.Headers['Retry-After'], [ref]$retryAfter) | Out-Null
                if ($retryAfter -and $retryAfter -gt $delay) { $delay = $retryAfter }
            }
            Write-Warning "Attempt $attempt failed: $($ex.Message). Retrying in ${delay}s..."
            Start-Sleep -Seconds $delay
        }
    }
}

# --- End moved functions ---

# Handle recipient email addresses - Azure Automation may pass as string or array
if ($SendGridRecipientEmailAddresses -is [string]) {
    # Convert comma-separated string to array
    $RecipientEmailArray = $SendGridRecipientEmailAddresses -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
} elseif ($SendGridRecipientEmailAddresses -is [array]) {
    # Already an array, just ensure trimmed and non-empty
    $RecipientEmailArray = $SendGridRecipientEmailAddresses | ForEach-Object { $_.Trim() } | Where-Object { $_ }
} else {
    # Convert to array for consistency
    $RecipientEmailArray = @($SendGridRecipientEmailAddresses | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

# Ensure RecipientEmailArray is properly initialized as an array
$RecipientEmailArray = @($RecipientEmailArray)
Write-Host "Recipients processed: $($RecipientEmailArray.Count) addresses" -ForegroundColor Green

# Coerce DisableAccounts to boolean. Azure Automation may pass it as a string.
$DisableAccountsBool = $false
if ($DisableAccounts -is [string]) {
    switch ($DisableAccounts.Trim().ToLower()) {
        '1' { $DisableAccountsBool = $true; break }
        'true' { $DisableAccountsBool = $true; break }
        'yes' { $DisableAccountsBool = $true; break }
        '0' { $DisableAccountsBool = $false; break }
        'false' { $DisableAccountsBool = $false; break }
        'no' { $DisableAccountsBool = $false; break }
        default { $DisableAccountsBool = $false }
    }
}
elseif ($DisableAccounts -is [bool]) { $DisableAccountsBool = $DisableAccounts }
elseif ($DisableAccounts -is [int]) { $DisableAccountsBool = ([int]$DisableAccounts -ne 0) }
else { $DisableAccountsBool = [bool]$DisableAccounts }
Write-Host "DisableAccounts coerced to boolean: $DisableAccountsBool" -ForegroundColor Green

# Disable-Leavers script - Process Workday leavers and disable accounts
# Resolve system temp directory and ensure it exists
$localtempdir = [System.IO.Path]::GetTempPath()
if (-not (Test-Path $localtempdir)) { New-Item -ItemType Directory -Path $localtempdir -Force | Out-Null }

# Timestamp once per run so CSVs are unique and not overwritten
$reportfilenamepattern = "Leavers_Report_-_Active_Directory_-_IT_withUPN"
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outputpath = Join-Path $localtempdir "${reportfilenamepattern}_${timestamp}.csv"

# Initialize Azure and Graph contexts (Managed Identity when available)
# Ensure required modules are present (best-effort installer)
Ensure-RequiredModules

# Initialize Azure and Graph contexts (Managed Identity when available)
Initialize-AzureContext
Initialize-GraphContext

# Build Azure Storage context using managed identity / connected account
try {
    $storagecontext = New-AzStorageContext -StorageAccountName $StorageAccountName -UseConnectedAccount -ErrorAction Stop
}
catch {
    throw "Failed to create storage context: $($_.Exception.Message)"
}
$inputpath = Join-Path $localtempdir (Split-Path -Leaf $BlobName)
Write-Host "Downloading input from Azure Storage: $StorageAccountName/$ContainerName/$BlobName -> $inputpath" -ForegroundColor Cyan
try {
    Get-AzStorageBlobContent -Container $ContainerName -Blob $BlobName -Destination $inputpath -Context $storagecontext -Force -ErrorAction Stop | Out-Null
}
catch {
    throw "Failed to download blob '$BlobName' from container '$ContainerName' in storage account '$StorageAccountName': $($_.Exception.Message)"
}

# Graph context is already initialized above via Initialize-GraphContext

Write-Host "Loading input CSV: $inputpath" -ForegroundColor Cyan
$rows = Import-Csv -Path $inputpath -ErrorAction Stop
if (-not $rows) {
    Write-Warning "No rows loaded from CSV. Exiting."
    return
}

Write-Host "Retrieving users from Entra ID (this may take time in large tenants)..." -ForegroundColor Cyan
# Single directory pull (avoid N x API calls). Adjust properties as needed.
$allusers = Invoke-WithRetry -Script {
    Get-MgUser -All -Property id, employeeId, mail, displayName, userPrincipalName, accountEnabled -ConsistencyLevel eventual -ErrorAction Stop |
    Select-Object id, userPrincipalName, employeeId, mail, displayName, accountEnabled
}

# Build fast lookup hash tables (case-insensitive keys)
$byemployeeid = @{}
$bymail = @{}
$bydisplay = @{}
foreach ($u in $allusers) {
    if ($u.employeeId -and -not $byemployeeid.ContainsKey($u.employeeId)) { $byemployeeid[$u.employeeId] = $u }
    if ($u.mail -and -not $bymail.ContainsKey($u.mail.ToLower())) { $bymail[$u.mail.ToLower()] = $u }
    if ($u.displayName) {
    $dnkey = $u.displayName.ToLower()
    if (-not $bydisplay.ContainsKey($dnkey)) { $bydisplay[$dnkey] = $u }
    }
}

$processed = 0
$total = $rows.Count
foreach ($row in $rows) {
    $processed++
    $employeeid = $row.'Employee ID'
    $email = $row.'Email - Primary Work'
    # Current file uses 'Legal Name' for display name source
    $displayname = $row.'Legal Name'

    # Ensure / (re)create output columns (idempotent with -Force)
    Add-Member -InputObject $row -NotePropertyName UPN             -NotePropertyValue $null -Force
    Add-Member -InputObject $row -NotePropertyName 'UPN-mail'      -NotePropertyValue $null -Force
    Add-Member -InputObject $row -NotePropertyName DisplayName_UPN -NotePropertyValue $null -Force
    Add-Member -InputObject $row -NotePropertyName AccountEnabled  -NotePropertyValue $null -Force
    Add-Member -InputObject $row -NotePropertyName MatchSource     -NotePropertyValue $null -Force
    Add-Member -InputObject $row -NotePropertyName OnPremDisabledActionResult -NotePropertyValue $null -Force

    $user = $null

    # 1. Match by employeeId (exact)
    if ($employeeid -and $byemployeeid.ContainsKey($employeeid)) {
        $user = $byemployeeid[$employeeid]
        $row.UPN = $user.userPrincipalName
        $row.AccountEnabled = $user.accountEnabled
        $row.MatchSource = 'EmployeeId'
    }
    # 2. Match by mail (case-insensitive exact)
    elseif ($email) {
        $lower = $email.ToLower()
        if ($bymail.ContainsKey($lower)) {
            $user = $bymail[$lower]
            $row.'UPN-mail' = $user.userPrincipalName
            $row.AccountEnabled = $user.accountEnabled
            $row.MatchSource = 'Mail'
        }
    }
    # 3. Match by display name (case-insensitive exact)
    if (-not $user -and $displayname) {
        $dnkey = $displayname.ToLower()
        if ($bydisplay.ContainsKey($dnkey)) {
            $user = $bydisplay[$dnkey]
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

if ($DisableAccountsBool) {
    Write-Host "Disabling matched enabled cloud accounts..." -ForegroundColor Yellow
    $matchedenabled = $rows | Where-Object { $_.MatchSource -and $_.AccountEnabled -eq $true }

    # Prepare on-prem AD context early if requested
    if (-not (Initialize-OnPremAD -Server $OnPremDomainController)) {
        Write-Warning "On-prem AD initialization failed. Only Entra ID accounts will be disabled."
        $onpremavailable = $false
    } else {
        $onpremavailable = $true
    }

    foreach ($r in $matchedenabled) {
        $upntodisable = $r.UPN
        if (-not $upntodisable) { $upntodisable = $r.'UPN-mail' }
        if (-not $upntodisable) { $upntodisable = $r.DisplayName_UPN }
        if (-not $upntodisable) { continue }
        try {
            Invoke-WithRetry -Script { Update-MgUser -UserId $upntodisable -AccountEnabled:$false -ErrorAction Stop } | Out-Null
            $r.DisabledActionResult = 'Disabled'
        }
        catch {
            $r.DisabledActionResult = "Error: $($_.Exception.Message)"
        }
    }

    # On-prem disable ALL matched accounts (even those already disabled in Entra)
    if ($onpremavailable) {
        Write-Host "Disabling matched accounts on-prem (all matched, regardless of cloud state)..." -ForegroundColor Yellow
        $matchedall = $rows | Where-Object { $_.MatchSource }
        foreach ($r in $matchedall) {
            $upntodisable = $r.UPN
            if (-not $upntodisable) { $upntodisable = $r.'UPN-mail' }
            if (-not $upntodisable) { $upntodisable = $r.DisplayName_UPN }
            if (-not $upntodisable) { continue }
            try {
                $aduserparams = @{ Identity = $upntodisable; Server = $OnPremDomainController; Properties = 'Enabled'; ErrorAction = 'Stop' }
                $cred = Get-OnPremAdCredential
                if ($cred) { $aduserparams['Credential'] = $cred }
                $aduser = $null
                $aduser = Get-ADUser @aduserparams
                if ($aduser.Enabled) {
                    $disparams = @{ Identity = $aduser.DistinguishedName; Server = $OnPremDomainController; ErrorAction = 'Stop' }
                    if ($cred) { $disparams['Credential'] = $cred }
                    Disable-ADAccount @disparams
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
$logpath = (Join-Path $localtempdir "Disable-Leavers_Report_$timestamp.log")
# Ensure output and log directories exist
$__pathstoensure = @((Split-Path -Parent $outputpath), (Split-Path -Parent $logpath))
foreach ($__d in $__pathstoensure) { if ($__d -and -not (Test-Path $__d)) { New-Item -ItemType Directory -Path $__d -Force | Out-Null } }
"Disable-Leavers run started: $(Get-Date)" | Out-File -FilePath $logpath -Encoding UTF8
"Input CSV: $inputpath" | Out-File -FilePath $logpath -Append
"Output CSV (will be written): $outputpath" | Out-File -FilePath $logpath -Append
"DisableAccounts parameter: $DisableAccounts (coerced: $DisableAccountsBool)" | Out-File -FilePath $logpath -Append

Write-Host "Writing output CSV: $outputpath" -ForegroundColor Cyan
$rows | Export-Csv -Path $outputpath -NoTypeInformation -ErrorAction Stop

# Build list of files to send (kept in temp directory)
$filestosend = @($inputpath, $outputpath, $logpath) | Where-Object { $_ -and (Test-Path $_) }
# Ensure $filestosend is an array
$filestosend = @($filestosend)
Write-Host "Files to send: $($filestosend.Count) files" -ForegroundColor Green

# Build report summary
$matched = $rows | Where-Object { $_.MatchSource }
$matchedcount = $matched.Count
$unknowncount = $rows.Count - $matchedcount
$bysource = $matched | Group-Object MatchSource | Select-Object Name, Count
$disabledsuccess = @($rows | Where-Object { $_.DisabledActionResult -eq 'Disabled' })
$disablederrors = @($rows | Where-Object { $_.DisabledActionResult -like 'Error:*' })
$onpremdisabled = @($rows | Where-Object { $_.OnPremDisabledActionResult -eq 'Disabled' })
$onpremdisableerr = @($rows | Where-Object { $_.OnPremDisabledActionResult -like 'Error:*' })
$disabledcount = $disabledsuccess.Count
$disableerrorcount = $disablederrors.Count

"" | Out-File -FilePath $logpath -Append
"Summary:" | Out-File -FilePath $logpath -Append
"Total rows:        $($rows.Count)" | Out-File -FilePath $logpath -Append
"Matched rows:      $matchedcount" | Out-File -FilePath $logpath -Append
foreach ($g in $bysource) { "  Matched by $($g.Name): $($g.Count)" | Out-File -FilePath $logpath -Append }
"Unknown rows:      $unknowncount" | Out-File -FilePath $logpath -Append
if ($DisableAccountsBool) {
    "Accounts disabled: $disabledcount" | Out-File -FilePath $logpath -Append
    "On-prem accounts disabled: $($onpremdisabled.Count)" | Out-File -FilePath $logpath -Append
    if ($disableerrorcount -gt 0) { "Disable errors:   $disableerrorcount" | Out-File -FilePath $logpath -Append }
    if ($onpremdisableerr.Count -gt 0) { "On-prem disable errors: $($onpremdisableerr.Count)" | Out-File -FilePath $logpath -Append }
}
else {
    "DisableAccounts is false - no accounts were disabled" | Out-File -FilePath $logpath -Append
}

if ($DisableAccountsBool -and $disableerrorcount -gt 0) {
    "" | Out-File -FilePath $logpath -Append
    "Disable error details:" | Out-File -FilePath $logpath -Append
    foreach ($e in $disablederrors) {
    $u = $e.UPN
    if (-not $u) { $u = $e.'UPN-mail' }
    if (-not $u) { $u = $e.DisplayName_UPN }
    "  Cloud: $u => $($e.DisabledActionResult)" | Out-File -FilePath $logpath -Append
    }
}
if ($DisableAccountsBool -and $onpremdisableerr.Count -gt 0) {
    "" | Out-File -FilePath $logpath -Append
    "On-prem disable error details:" | Out-File -FilePath $logpath -Append
    foreach ($e in $onpremdisableerr) {
    $u = $e.UPN
    if (-not $u) { $u = $e.'UPN-mail' }
    if (-not $u) { $u = $e.DisplayName_UPN }
    "  OnPrem: $u => $($e.OnPremDisabledActionResult)" | Out-File -FilePath $logpath -Append
    }
}

"Run completed: $(Get-Date)" | Out-File -FilePath $logpath -Append
Write-Host "Report log written: $logpath" -ForegroundColor Green


