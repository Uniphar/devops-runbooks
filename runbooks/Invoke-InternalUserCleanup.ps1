<#
.SYNOPSIS
Disable inactive user accounts in Microsoft Entra ID (Azure AD) and on-premises Active Directory

.DESCRIPTION
This script finds tenant users (Members) that have an employeeId set and evaluates their sign-in activity
to determine whether they should be notified or disabled based on inactivity thresholds.

Key behavior:
- Only processes enabled tenant users of type 'Member' that have the `employeeId` attribute defined.
- Uses Microsoft Graph (beta) to read sign-in activity and manager information.
- Compares sign-in activity with on-premises AD LastLogonDate to ensure user is inactive in both systems.
- Sends notification emails before disabling and disables accounts after the configured inactivity period.
- When `$testing` is `$true` the script will not perform disabling actions and will log email activity locally.

.PARAMETER sendGridApiKeyKvName
Name of the Key Vault that contains secrets. This script requires the Key Vault name to be supplied via
the `-sendGridApiKeyKvName` parameter (no Automation variable fallback).

.PARAMETER sendGridApiKeyKvSecretName
Name of the Key Vault secret that contains the SendGrid API key. Required when running with `-testing:$false`.

.PARAMETER onPremKeyVaultName
Optional Key Vault name to read on-prem AD credentials from. If not supplied, the SendGrid vault is used.

.PARAMETER domainAdminUserSecretName
Name of the Key Vault secret that contains the on-prem AD service account username.

.PARAMETER domainAdminPwdSecretName
Name of the Key Vault secret that contains the on-prem AD service account password.

.PARAMETER domainControllerName
FQDN or hostname of the on-premises Active Directory domain controller to query.

.PARAMETER sendGridSenderEmailAddress
Sender email address used when sending messages via SendGrid.

.PARAMETER sendGridRecipientEmailAddresses
Recipient addresses for administrative reports (array). The script also sends notifications to each user and their manager.

.PARAMETER inactivityTime
Number of days of inactivity after which a user will be disabled (default: 45).

.PARAMETER userWarningThreshold
Number of days of inactivity when a warning notification should be sent (default: 35).

.PARAMETER groupId
Exclusion group id (GUID). Members of this group (including nested group members) are excluded from evaluation.

.PARAMETER testing
Switch for testing mode. When `$true` no disabling actions are performed and emails are logged to a local file.

.NOTES
Requires:
 - Az.KeyVault module (Get-AzKeyVaultSecret)
 - ActiveDirectory module (Get-ADUser / Disable-ADAccount) on the Hybrid Worker
 - Microsoft.Graph.Authentication and Microsoft.Graph.* modules for SignInActivity and user operations
 - Automation account managed identity with Key Vault and Graph permissions
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string] $sendGridApiKeyKvName,
    [Parameter(Mandatory = $false)]
    [string] $sendGridApiKeyKvSecretName,
    [Parameter(Mandatory = $false)]
    [string] $onPremKeyVaultName,
    [Parameter(Mandatory = $false)]
    [string] $domainAdminUserSecretName,
    [Parameter(Mandatory = $false)]
    [string] $domainAdminPwdSecretName,
    [Parameter(Mandatory = $true)]
    [string] $domainControllerName,
    [Parameter(Mandatory = $false)]
    [string] $sendGridSenderEmailAddress,
    [Parameter(Mandatory = $false)]
    [object] $sendGridRecipientEmailAddresses,
    [Parameter(Mandatory = $false)]
    [int] $inactivityTime = 45,
    [Parameter(Mandatory = $false)]
    [int] $userWarningThreshold = 35,
    [Parameter(Mandatory = $true)]
    [string] $groupId,
    [Parameter(Mandatory = $false)]
    [bool] $testing = $true
)


# Import required modules
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Import-Module Az.KeyVault -ErrorAction Stop
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Beta.Users -ErrorAction Stop
    Import-Module Microsoft.Graph.Beta.Identity.DirectoryManagement -ErrorAction Stop
    Import-Module Microsoft.Graph.Groups -ErrorAction Stop
    
    Write-Output "All required modules imported successfully"
}
catch {
    Write-Error "Failed to import required modules: $_"
    return
}



Connect-AzAccount -Identity -ErrorAction Stop
    Write-Output "Successfully connected to Azure using managed identity"
# Connect to Microsoft Graph using the runbook managed identity.
# Note: Permissions are assigned to the managed identity in Azure AD. Required Graph permissions (examples):
#   - User.Read.All, AuditLog.Read.All, Group.Read.All for read operations
#   - User.ReadWrite.All if this runbook will update/disable users in production

Connect-MgGraph -Identity -NoWelcome -ErrorAction Stop
     Write-Output "Successfully connected to Microsoft Graph using managed identity"

# Normalize administrative recipient input so Azure Automation single-value bindings do not break array expectations.
if ($null -eq $sendGridRecipientEmailAddresses) {
    $sendGridRecipientEmailAddresses = @()
}
elseif ($sendGridRecipientEmailAddresses -is [System.Collections.IEnumerable] -and $sendGridRecipientEmailAddresses -isnot [string]) {
    $sendGridRecipientEmailAddresses = @($sendGridRecipientEmailAddresses | ForEach-Object { if ($_ -ne $null) { $_.ToString().Trim() } } | Where-Object { $_ })
}
else {
    $rawRecipients = $sendGridRecipientEmailAddresses.ToString().Trim()

    if ($rawRecipients.StartsWith('[') -and $rawRecipients.EndsWith(']')) {
        try {
            $parsedRecipients = $rawRecipients | ConvertFrom-Json -ErrorAction Stop
            $sendGridRecipientEmailAddresses = @($parsedRecipients | ForEach-Object { if ($_ -ne $null) { $_.ToString().Trim() } } | Where-Object { $_ })
        }
        catch {
            $sendGridRecipientEmailAddresses = @($rawRecipients -split "[,;`r`n]" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        }
    }
    else {
        $sendGridRecipientEmailAddresses = @($rawRecipients -split "[,;`r`n]" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }
}

$sendGridRecipientEmailAddresses = [string[]]$sendGridRecipientEmailAddresses

# Determine Key Vault name: the runbook requires the Key Vault name to be supplied via
# the `-sendGridApiKeyKvName` parameter. The script will error and return if it's not provided.
if ($sendGridApiKeyKvName) {
    $vaultName = $sendGridApiKeyKvName
}
else {
    Write-Error "Parameter -sendGridApiKeyKvName is required so the runbook can retrieve secrets from Key Vault."
    return
}

$sendGridApiKey = ''

if ($sendGridApiKeyKvSecretName) {
    try {
        $sendGridSecret = Get-AzKeyVaultSecret -VaultName $vaultName -Name $sendGridApiKeyKvSecretName -ErrorAction Stop
        if ($sendGridSecret.SecretValueText) {
            $sendGridApiKey = $sendGridSecret.SecretValueText
        }
        else {
            $sendGridApiKey = [Net.NetworkCredential]::new('', $sendGridSecret.SecretValue).Password
        }
    }
    catch {
        Write-Error "Failed to retrieve SendGrid API key '$sendGridApiKeyKvSecretName' from Key Vault '$vaultName'. $_"
        return
    }
}
elseif (-not $testing) {
    Write-Error 'When -testing is set to $false you must supply -sendGridApiKeyKvSecretName so the script can send email notifications.'
    return
}
# Determine which Key Vault to use for on-prem credentials (prefer explicit parameter)
$onPremVault = if ($onPremKeyVaultName) { $onPremKeyVaultName } else { $vaultName }

# Retrieve on-premises AD domain admin username/password from Key Vault (as SecureString)
$domainUserSecret = Get-AzKeyVaultSecret -VaultName $onPremVault -Name $domainAdminUserSecretName
$domainPwdSecret = Get-AzKeyVaultSecret -VaultName $onPremVault -Name $domainAdminPwdSecretName

if (-not $domainUserSecret -or -not $domainPwdSecret) {
    Write-Error "Could not retrieve on-prem credentials from Key Vault '$onPremVault'. Ensure secrets exist: $domainAdminUserSecretName, $domainAdminPwdSecretName"
    return
}

# Get-AzKeyVaultSecret returns a PSSecretRecord; .SecretValue is a SecureString. Use it directly for PSCredential.
$domainUser = [Net.NetworkCredential]::new('', $domainUserSecret.SecretValue).Password
$domainPasswordSecure = $domainPwdSecret.SecretValue
$adCredentials = New-Object System.Management.Automation.PSCredential ($domainUser, $domainPasswordSecure)

# Domain controller/server name is supplied via parameter `$domainControllerName`
$domainController = $domainControllerName

# Calculate days remaining before disabling (define early, used in emails)
$daysRemaining = $inactivityTime - $userWarningThreshold

# Calculate the mid-point threshold for second notification
$midPointThreshold = [math]::Round(($userWarningThreshold + $inactivityTime) / 2)
 
# Directory where temporary CSV reports are written (Automation runbook uses $env:TEMP)
$reportDir = $env:TEMP

# Hardcoded SendGrid API endpoint
$sendGridApiEndpoint = 'https://api.sendgrid.com/v3/mail/send'

# Guard: ensure runbook is running under Windows PowerShell (5.1) when using RSAT/AD cmdlets
# If the Automation runbook is configured to use PowerShell 7 (Core) but the Hybrid Worker
# does not have the pwsh executable, Azure will fail to start the job with an unclear error.
if ($PSVersionTable.PSEdition -eq 'Core') {
    Write-Error "This runbook requires Windows PowerShell (Desktop/5.1) because it uses the ActiveDirectory RSAT module."
    Write-Error "Options: (a) Change the runbook runtime to 'Windows PowerShell (5.1)' in the Automation runbook settings, or (b) install PowerShell 7 (pwsh) on the Hybrid Worker and ensure it's in PATH."
    return
}

# Function to disable user in on-premises AD
function Disable-OnPremADUser {
    param (
        [string]$userPrincipalName
    )
    try {
        $user = Get-ADUser -Server $domainController -Credential $adCredentials -Filter { UserPrincipalName -eq $userPrincipalName } -Properties Description -ErrorAction Stop
        if ($user) {
            # Disable the account
            Disable-ADAccount -Server $domainController -Credential $adCredentials -Identity $user -ErrorAction Stop
            
            # Update the Description field - prepend "INACTIVE-USER-DISABLED" if not already present
            $currentDescription = $user.Description
            if (-not $currentDescription) {
                $newDescription = "INACTIVE-USER-DISABLED"
            }
            elseif (-not $currentDescription.StartsWith("INACTIVE-USER-DISABLED")) {
                $newDescription = "INACTIVE-USER-DISABLED $currentDescription"
            }
            else {
                $newDescription = $currentDescription # Already marked
            }
            
            Set-ADUser -Server $domainController -Credential $adCredentials -Identity $user -Description $newDescription -ErrorAction Stop
            
            Write-Output "Disabled on-prem AD account for user: $userPrincipalName (Description updated)"
            return "Success"
        }
        else {
            Write-Warning "User not found in on-prem AD: $userPrincipalName"
            return "User not found"
        }
    }
    catch {
        Write-Error "Failed to disable on-prem AD account for user '$userPrincipalName': $_"
        return "Error: $_"
    }
}

# Function to disable user in Azure AD via Microsoft Graph PowerShell module
function Disable-MgUser {
    param (
        [string]$userId
    )
    try {
        Update-MgUser -UserId $userId -AccountEnabled:$false
        Write-Debug "Disabled Azure AD account for user: $userId"
        return "Success"
    }
    catch {
        Write-Error "Failed to disable Azure AD account for user: $userId"
        return "Failed"
    }
}

function Get-GroupMembers {
    # Recursive function to get all members of a group (includes nested groups)
    param (
        [string]$groupId,
        [ref]$exclusion
    )

    if (-not $groupId) {
        Write-Error "groupId cannot be empty."
        return
    }

    try {
        Write-Output "Retrieving members for group ID: $groupId"
        
        # Try to get the group first to validate it exists
        $group = Get-MgGroup -GroupId $groupId -ErrorAction Stop
        Write-Output "Found group: $($group.DisplayName)"
        
        # Get group members using the correct cmdlet
        $members = Get-MgGroupMember -GroupId $groupId -All -ErrorAction Stop
        Write-Output "Retrieved $($members.Count) members from group"

        foreach ($member in $members) {
            $memberType = $member.AdditionalProperties.'@odata.type'
            
            if (-not $memberType -and $member.PSObject.Properties['@odata.type']) {
                $memberType = $member.PSObject.Properties['@odata.type'].Value
            }

            if ($memberType -eq "#microsoft.graph.user") {
                $userPrincipalName = $null

                if ($member.AdditionalProperties -and $member.AdditionalProperties.ContainsKey('userPrincipalName')) {
                    $userPrincipalName = $member.AdditionalProperties['userPrincipalName']
                }
                elseif ($member.PSObject.Properties['UserPrincipalName']) {
                    $userPrincipalName = $member.UserPrincipalName
                }
                else {
                    # Fall back to an explicit lookup if the lightweight member payload lacks UPN
                    $user = Get-MgUser -UserId $member.Id -ErrorAction SilentlyContinue
                    if ($user) {
                        $userPrincipalName = $user.UserPrincipalName
                    }
                }

                if ($userPrincipalName) {
                    [void]$exclusion.Value.Add($userPrincipalName)
                    # Remove individual user output - will be shown in summary
                }
                else {
                    Write-Warning "Could not resolve user principal name for member '$($member.Id)'; skipping."
                }
            }
            elseif ($memberType -eq "#microsoft.graph.group") {
                # Get the nested group and recursively process its members
                $nestedGroup = Get-MgGroup -GroupId $member.Id -ErrorAction SilentlyContinue
                if ($nestedGroup) {
                    Write-Output "Processing nested group: $($nestedGroup.DisplayName)"
                    Get-GroupMembers -GroupId $nestedGroup.Id -Exclusion $exclusion
                }
            }
        }
    }
    catch {
        Write-Error "Failed to retrieve group members for group '$groupId': $_"
        Write-Error "Error details: $($_.Exception.Message)"
        if ($_.Exception.InnerException) {
            Write-Error "Inner exception: $($_.Exception.InnerException.Message)"
        }
    }
}


# Function to send an email using SendGrid, with optional attachments.
function Send-Email {
    <#
.SYNOPSIS
Sends an email using SendGrid, with optional attachments.

.DESCRIPTION
This function sends an email with the specified content to the given recipient email addresses using SendGrid's API. Attachments are optional.

.PARAMETER sendGridApiKey
The API key for authenticating with SendGrid.

.PARAMETER senderEmailAddress
The email address of the sender.

.PARAMETER recipientEmailAddresses
An array of recipient email addresses.

.PARAMETER subject
The subject of the email.

.PARAMETER content
The plain-text content of the email.

.PARAMETER attachments
An optional array of attachment objects. Each object must include:
  - file : path to the file to attach
  - type : MIME type string (e.g. "text/csv")
#>
    Param(
        [Parameter(Mandatory = $true)]
        [string] $sendGridApiKey,

        [Parameter(Mandatory = $true)]
        [string] $senderEmailAddress,

        [Parameter(Mandatory = $true)]
        [string[]] $recipientEmailAddresses,

        [Parameter(Mandatory = $true)]
        [String] $subject,

        [Parameter(Mandatory = $true)]
        [String] $content,

        [Parameter(Mandatory = $false)]
        [Object[]] $attachments
    )

    # Log email details when testing
    if ($testing) {
        $date = Get-Date -Format "yyyy-MM-dd"
        $timestamp = Get-Date -Format "HHmmss"
        $logFile = "$env:TEMP\email_log_${date}_${timestamp}.txt"
        # Use parentheses to ensure the pipeline is evaluated before -join, and use the 'file' property.
        $attachmentNames = if ($attachments) { ($attachments | ForEach-Object { [System.IO.Path]::GetFileName($_.file) }) -join ', ' } else { '' }
        $logContent = "To: $($recipientEmailAddresses -join ', ')" + [Environment]::NewLine +
        "Subject: $subject" + [Environment]::NewLine +
        "Content: $content" + [Environment]::NewLine +
        "Attachments: $attachmentNames" + [Environment]::NewLine +
        "----------------------------------------" + [Environment]::NewLine
        
        try {
            Add-Content -Path $logFile -Value $logContent -ErrorAction Stop
            Write-Output "Email logged to: $logFile"
        }
        catch {
            Write-Warning "Could not write to log file '$logFile'. Error: $_"
        }
    }
    
    # Send email via SendGrid
    try {
        $headers = @{
            "Authorization" = "Bearer $sendGridApiKey"
            "Content-Type"  = "application/json"
        }

        $attachmentObjects = if ($attachments) {
            $attachments | ForEach-Object {
                $contentCsv = Get-Content $_.file -Raw

                @{
                    content     = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($contentCsv))
                    filename    = [System.IO.Path]::GetFileName($_.file)
                    type        = $_.type
                    disposition = "attachment"
                }
            }
        }
        else {
            @()
        }

        $body = @{
            from             = @{ email = $senderEmailAddress }
            personalizations = @(@{ to = @($recipientEmailAddresses | ForEach-Object { @{ email = $_ } }) })
            subject          = $subject
            content          = @(@{ type = "text/plain"; value = $content })
            attachments      = $attachmentObjects
        }

        $bodyJson = $body | ConvertTo-Json -Depth 4
        Invoke-RestMethod -Uri $sendGridApiEndpoint -Method Post -Headers $headers -Body $bodyJson -ErrorAction Stop
        Write-Output "Email sent successfully to: $($recipientEmailAddresses -join ', ')"
    }
    catch {
        Write-Error "Failed to send email via SendGrid: $_"
    }
}


# Get all exclusion group members' UPNs
$exclusion = [System.Collections.ArrayList]@()
Write-Output "Processing exclusion group with ID: $groupId"
Get-GroupMembers -GroupId $groupId -Exclusion ([ref]$exclusion)
Write-Output "Exclusion list populated with $($exclusion.Count) users"
if ($exclusion.Count -gt 0) {
    # Show only the first 10 exclusion UPNs
    $previewCount = [Math]::Min(10, $exclusion.Count)
    $preview = $exclusion[0..($previewCount - 1)] -join ', '
    if ($exclusion.Count -le 10) {
        Write-Output "Exclusion list (all $($exclusion.Count) users): $preview"
    }
    else {
        Write-Output "Exclusion list (showing first 10 of $($exclusion.Count) users): $preview"
    }
}

# Gather all users in tenant (only users with employeeID defined)
Write-Output "Retrieving all users from Microsoft Graph (Beta)..."
$allUsers = Get-MgBetaUser -Property SignInActivity,EmployeeId,AccountEnabled,UserType,DisplayName,UserPrincipalName,Mail,CompanyName,CreatedDateTime,Id -All | Where-Object { $_.AccountEnabled -and $_.UserType -eq "Member" -and $_.EmployeeId }

# Prepare on-prem Active Directory activity lists for two cutoff dates:
#  - $cutoffDate (for disabling decisions)
#  - $cutoffDate2 (for warning/notification decisions)
$activeUPNs = $null
# Calculate the cutoff date
$cutoffDate = (Get-Date).AddDays(-$inactivityTime)
$cutoffDate2 = (Get-Date).AddDays(-$userWarningThreshold)

# Get list of all UPNs from on-prem AD that were active within the inactivity time
Write-Output "Connecting to on-premises AD domain controller: $domainController"
try {
    $activeUsers = Get-ADUser -Server $domainController -Credential $adCredentials -Filter { LastLogonDate -ge $cutoffDate } -Properties UserPrincipalName, LastLogonDate -ErrorAction Stop
    $activeUsers2 = Get-ADUser -Server $domainController -Credential $adCredentials -Filter { LastLogonDate -ge $cutoffDate2 } -Properties UserPrincipalName, LastLogonDate -ErrorAction Stop
    Write-Output "Successfully retrieved on-prem AD user data. Active users (cutoff $cutoffDate): $($activeUsers.Count), Active users (cutoff $cutoffDate2): $($activeUsers2.Count)"
}
catch {
    Write-Error "Failed to contact on-premises AD domain controller '$domainController'. Error: $_"
    Write-Error "Verify: (1) Domain controller name is correct, (2) Hybrid Worker can reach the DC, (3) Credentials are valid, (4) Network/firewall allows LDAP traffic."
    return
}

# Extract UPNs
$activeUPNs = $activeUsers | Select-Object -ExpandProperty UserPrincipalName
$activeUPNs2 = $activeUsers2 | Select-Object -ExpandProperty UserPrincipalName

# Create a new empty array list object for disabling
# Create a new empty array list object for disabling
$report = [System.Collections.Generic.List[Object]]::new()
# Create a new empty array list object for notification
$notification = [System.Collections.Generic.List[Object]]::new()
# Create a new empty array list object for mid-point notification
$midPointNotification = [System.Collections.Generic.List[Object]]::new()
   
Foreach ($user in $allUsers) {
    # Null variables (initialize variables that are used later)
    $licenses = $null
    $manager = $null
    $maxDate = $null
       
    
    # Display progress output 
    Write-Debug "Gathering sign-in information for $($user.DisplayName)"


    # Count the last signing date from all posible variants
    # Retrieve the date values
    $lastInteractiveSignIn = $user.SignInActivity.LastSignInDateTime
    $lastNonInteractiveSignIn = $user.SignInActivity.LastNonInteractiveSignInDateTime
    $lastSuccessfulSignInDate = $user.SignInActivity.LastSuccessfulSignInDateTime

    $maxDate = $null

    if ($null -ne $lastInteractiveSignIn) {
        $maxDate = $lastInteractiveSignIn
    }

    if ($null -ne $lastNonInteractiveSignIn -and ($null -eq $maxDate -or $lastNonInteractiveSignIn -gt $maxDate)) {
        $maxDate = $lastNonInteractiveSignIn
    }

    if ($null -ne $lastSuccessfulSignInDate -and ($null -eq $maxDate -or $lastSuccessfulSignInDate -gt $maxDate)) {
        $maxDate = $lastSuccessfulSignInDate
    }
    #reporting on the screen and get $daysInactive for user
    if ($maxDate) {
        $daysInactive = [math]::Round(((Get-Date) - $maxDate).TotalDays)
        Write-Debug "Last sign in date is $($maxDate), Days of inactivity: $daysInactive"
    }
    else {
        Write-Debug "Last sign in date is not available"
    }
    # Retrieve account creation date
    $accountCreationDate = $user.CreatedDateTime
    if (-not $accountCreationDate) {
        $accountCreationDate = [datetime]"1/1/2000" #put some old date if it is empty
    }
    $daysSinceCreation = [math]::Round(((Get-Date) - $accountCreationDate).TotalDays)

    if ($daysSinceCreation -gt 21 -and $null -ne $maxDate -and $maxDate -lt (Get-Date).AddDays(-$inactivityTime)) {
        # if user inactive and the account is not new, then 
        # Get current user license information
        try {
            $licenses = (Get-MgBetaUserLicenseDetail -UserId $user.id -ErrorAction SilentlyContinue).SkuPartNumber -join ", "
        }
        catch {
            $licenses = "Error retrieving licenses"
        }
    
        # Proceed silently to get manager information
        try {
            $managerid = Get-MgUserManager -UserId $user.id -ErrorAction SilentlyContinue

            if ($managerid) {
                $manager = get-mguser -UserId $managerid.Id -ErrorAction SilentlyContinue
            }
        }
        catch {
            $manager = $null
        }
    
        # Verify if the user is in the $activeUPNs list, if not, continue
        if ($user.UserPrincipalName -notin $activeUPNs) {
    
            # Verify if the user is in the $exclusion list, if not, continue
            if ($user.UserPrincipalName -notin $exclusion) {
    
    
                # Create informational object to add to report
                $obj1 = [pscustomobject][ordered]@{
                    DisplayName       = $user.DisplayName
                    UserPrincipalName = $user.UserPrincipalName
                    Email             = $user.Mail
                    Manager           = $manager.DisplayName
                    ManagerEmail      = $manager.Mail
                    Licenses          = $licenses
                    Company           = $user.CompanyName
                    CreatedDateTime   = $accountCreationDate
                    LastActivityDate  = $maxDate
                }
                Write-Debug "Adding user to disabling list $($user.DisplayName)"
                # Add current user info to report
                $report.Add($obj1)
            }
        }
    }



    #create a report for users to send FIRST warning (at userWarningThreshold - e.g., 35 days)
    if ($daysInactive -eq $userWarningThreshold) {
        # if user inactive then 
        
        # Get current user license information
        try {
            $licenses = (Get-MgBetaUserLicenseDetail -UserId $user.id -ErrorAction SilentlyContinue).SkuPartNumber -join ", "
        }
        catch {
            $licenses = "Error retrieving licenses"
        }
            
        # Proceed silently to get manager information
        try {
            $managerid = Get-MgUserManager -UserId $user.id -ErrorAction SilentlyContinue

            if ($managerid) {
                $manager = get-mguser -UserId $managerid.Id -ErrorAction SilentlyContinue
            }
        }
        catch {
            $manager = $null
        }
    
        # Verify if the user is in the $activeUPNs list, if not, continue
        if ($user.UserPrincipalName -notin $activeUPNs2) {
    
            # Verify if the user is in the $exclusion list, if not, continue
            if ($user.UserPrincipalName -notin $exclusion) {
    
                # Create informational object to add to report
                $obj2 = [pscustomobject][ordered]@{
                    DisplayName       = $user.DisplayName
                    UserPrincipalName = $user.UserPrincipalName
                    Email             = $user.Mail
                    Manager           = $manager.DisplayName
                    ManagerEmail      = $manager.Mail
                    Licenses          = $licenses
                    Company           = $user.CompanyName
                    CreatedDateTime   = $accountCreationDate
                    LastActivityDate  = $maxDate
                }
                Write-Debug "Adding user to notification list (first warning) $($user.DisplayName)"
                # Add current user info to report
                $notification.Add($obj2)
            }
        }
    }

    #create a report for users to send SECOND warning (at mid-point - e.g., 40 days)
    if ($daysInactive -eq $midPointThreshold) {
        # if user inactive then 
        
        # Get current user license information
        try {
            $licenses = (Get-MgBetaUserLicenseDetail -UserId $user.id -ErrorAction SilentlyContinue).SkuPartNumber -join ", "
        }
        catch {
            $licenses = "Error retrieving licenses"
        }
            
        # Proceed silently to get manager information
        try {
            $managerid = Get-MgUserManager -UserId $user.id -ErrorAction SilentlyContinue

            if ($managerid) {
                $manager = get-mguser -UserId $managerid.Id -ErrorAction SilentlyContinue
            }
        }
        catch {
            $manager = $null
        }
    
        # Verify if the user is in the $activeUPNs list, if not, continue
        if ($user.UserPrincipalName -notin $activeUPNs2) {
    
            # Verify if the user is in the $exclusion list, if not, continue
            if ($user.UserPrincipalName -notin $exclusion) {
    
                # Create informational object to add to report
                $obj3 = [pscustomobject][ordered]@{
                    DisplayName       = $user.DisplayName
                    UserPrincipalName = $user.UserPrincipalName
                    Email             = $user.Mail
                    Manager           = $manager.DisplayName
                    ManagerEmail      = $manager.Mail
                    Licenses          = $licenses
                    Company           = $user.CompanyName
                    CreatedDateTime   = $accountCreationDate
                    LastActivityDate  = $maxDate
                }
                Write-Debug "Adding user to mid-point notification list (second warning) $($user.DisplayName)"
                # Add current user info to report
                $midPointNotification.Add($obj3)
            }
        }
    }

}

$report | Export-CSV -path "$reportDir\disabled_users.csv" -NoTypeInformation
$notification | Export-CSV -path "$reportDir\notification_list.csv" -NoTypeInformation
$midPointNotification | Export-CSV -path "$reportDir\midpoint_notification_list.csv" -NoTypeInformation

# Initialize report array
$disableReport = @()

# Iterate over the report and disable users
foreach ($user in $report) {
    if (-not $Testing) {
        $onPremResult = Disable-OnPremADUser -userPrincipalName $user.UserPrincipalName
        $azureResult = Disable-MgUser -userId $user.UserPrincipalName
    }
    else {
        $onPremResult = "Testing mode - no action taken"
        $azureResult = "Testing mode - no action taken"
    }

    # Add result to report
    $disableReport += [PSCustomObject]@{
        UserPrincipalName = $user.UserPrincipalName
        OnPremResult      = $onPremResult
        AzureResult       = $azureResult
    }
}

# Export the report to CSV
$disableReport | Export-Csv -Path "$reportDir\DisableReport.csv" -NoTypeInformation

# Send notifications to managers about disabled accounts (THIRD notification)
foreach ($user in $report) {
    if (-not $testing) {
        # Only notify if we have a manager email and the disable was successful
        if ($user.ManagerEmail) {
            $disableStatus = $disableReport | Where-Object { $_.UserPrincipalName -eq $user.UserPrincipalName }
            
            $managerEmailSubject = "Account Disabled: $($user.DisplayName)"
            $managerEmailContent = @"
Dear $($user.Manager),

This notification confirms that the user account for $($user.DisplayName) has been disabled due to inactivity.

ACCOUNT DETAILS:
- User: $($user.DisplayName) ($($user.Email))
- Last Activity: $($user.LastActivityDate)
- Disabled Date: $(Get-Date -Format "yyyy-MM-dd")
- Reason: No sign-in activity for $inactivityTime days

DISABLE STATUS:
- Azure AD: $($disableStatus.AzureResult)
- On-Premises AD: $($disableStatus.OnPremResult)

PREVIOUS NOTIFICATIONS:
This is the third notification regarding this account:
- First notification was sent at $userWarningThreshold days of inactivity
- Second notification was sent at $midPointThreshold days of inactivity
- Account has now been disabled after $inactivityTime days of inactivity

NEXT STEPS:
If this user is no longer with the organization or no longer needs this account:
- No action required - the account will remain disabled
- IT will proceed with license removal and full deprovisioning

If this user still requires access:
The user can request account reactivation through the IT Service Portal

This is an automated notification. Do not reply to this email.
"@
            
            Send-Email -sendGridApiKey $sendGridApiKey `
                -senderEmailAddress $sendGridSenderEmailAddress `
                -recipientEmailAddresses @($user.ManagerEmail) `
                -subject $managerEmailSubject `
                -content $managerEmailContent
            
            Write-Output "Sent disable notification (3rd) to manager: $($user.ManagerEmail) for user: $($user.DisplayName)"
        }
        else {
            Write-Warning "No manager email for disabled user: $($user.DisplayName). Skipping manager notification."
        }
    }
    else {
        Write-Output "Testing mode: Would send disable notification (3rd) to manager for user $($user.DisplayName)"
    }
}

# Send report after disabling to IT support teams
$attachments = @(
    @{
        file = "$reportDir\DisableReport.csv"
        type = "text/csv"
    }
    @{
        file = "$reportDir\disabled_users.csv"
        type = "text/csv"
    }
    @{
        file = "$reportDir\notification_list.csv"
        type = "text/csv"
    }
    @{
        file = "$reportDir\midpoint_notification_list.csv"
        type = "text/csv"
    }
)

$emailSubject = "Inactive Users Report - Account Disabling Summary"
$emailContent = @"
Dear IT Operations Team,

This is an automated report from the Inactive User Management system.

SUMMARY:
- Total users disabled: $($report.count)
- Total users notified (first warning - $userWarningThreshold days): $($notification.count)
- Total users notified (second warning - $midPointThreshold days): $($midPointNotification.count)

The attached CSV files contain detailed information about:
1. DisableReport.csv - Results of disable operations (success/failure status)
2. disabled_users.csv - Users that were disabled in this run
3. notification_list.csv - Users that received first warning ($userWarningThreshold days)
4. midpoint_notification_list.csv - Users that received second warning ($midPointThreshold days)

PROCESS OVERVIEW:
Accounts are disabled after $inactivityTime days of inactivity (no sign-in activity in Microsoft Entra ID or on-premises Active Directory). 

NOTIFICATION STAGES:
1. First notification: After $userWarningThreshold days of inactivity ($daysRemaining days until disabling)
2. Second notification: After $midPointThreshold days of inactivity (reminder to managers)
3. Third notification: After $inactivityTime days - account disabled (confirmation to managers)

NEXT STEPS:
- Review the attached reports for accuracy
- Contact users or their managers if you have questions about specific accounts
- Users can request account reactivation through the IT Service Portal
- if you are sure users no longer need their accounts, please proceed with full deprovisioning process and license removal.

Please review the attached reports and take any necessary follow-up actions.
"@

Send-Email -sendGridApiKey $sendGridApiKey `
    -senderEmailAddress $sendGridSenderEmailAddress `
    -recipientEmailAddresses $sendGridRecipientEmailAddresses `
    -subject $emailSubject `
    -content $emailContent `
    -attachments $attachments

# Iterate over the report and send FIRST notification to user and manager
foreach ($user in $notification) {
    if (-not $testing) {
        # Validate that user has an email address
        if (-not $user.Email) {
            Write-Warning "User $($user.DisplayName) has no email address. Skipping notification."
            continue
        }
        
        # Build recipient list (user + manager if manager has email)
        $userRecipients = @($user.Email)
        if ($user.ManagerEmail) {
            $userRecipients += $user.ManagerEmail
        }
        else {
            Write-Warning "User $($user.DisplayName) has no manager email. Notification will only be sent to user."
        }
        
        $userEmailSubject = "Action Required: Your Account Will Be Disabled in $daysRemaining Days (First Notice)"
        $userEmailContent = @"
Dear $($user.DisplayName),

We have noticed that your user account has been inactive for $userWarningThreshold days. According to our account security policy, inactive accounts are disabled after $inactivityTime days of inactivity.

YOUR ACCOUNT STATUS:
- Current Status: Enabled (but inactive)
- Last Activity: More than $userWarningThreshold days ago
- Days Until Disabling: $daysRemaining

ACTION REQUIRED:
To prevent your account from being disabled, please log in to your account as soon as possible. A single successful login will reset your activity status and prevent the account from being disabled.

HOW TO RE-ACTIVATE:
1. Sign in to Microsoft 365 or your on-premises systems
2. Enter your username and password
3. Complete any multi-factor authentication prompts

NOTIFICATION SCHEDULE:
- This is the FIRST notification (at $userWarningThreshold days of inactivity)
- You will receive a second reminder in approximately $([math]::Round($midPointThreshold - $userWarningThreshold)) days
- Your account will be disabled after $inactivityTime days of total inactivity

MANAGER NOTIFICATION:
Your manager ($($user.Manager)) has also been notified of this status.

NEED HELP?
If you have questions or need to request an exemption, please:
Create a support request via the IT Help Desk

Please note: This is an automated notification. Do not reply to this email.
"@
        
        Send-Email -sendGridApiKey $sendGridApiKey `
            -senderEmailAddress $sendGridSenderEmailAddress `
            -recipientEmailAddresses $userRecipients `
            -subject $userEmailSubject `
            -content $userEmailContent
        
        Write-Output "Sent first warning to user and manager: $($user.DisplayName)"
    }
    else {
        Write-Output "Testing mode: Skipping first notification email to user $($user.DisplayName) and manager"
    }
}

# Iterate over the mid-point report and send SECOND notification to manager only
foreach ($user in $midPointNotification) {
    if (-not $testing) {
        # Only notify if we have a manager email
        if ($user.ManagerEmail) {
            $daysUntilDisable = $inactivityTime - $midPointThreshold
            
            $managerEmailSubject = "Reminder: $($user.DisplayName)'s Account Will Be Disabled in $daysUntilDisable Days (Second Notice)"
            $managerEmailContent = @"
Dear $($user.Manager),

This is a follow-up notification regarding the inactive account for your direct report, $($user.DisplayName).

ACCOUNT STATUS UPDATE:
- User: $($user.DisplayName) ($($user.Email))
- Current Status: Still Enabled (but inactive)
- Days of Inactivity: $midPointThreshold days
- Days Until Disabling: $daysUntilDisable days
- Last Activity: $($user.LastActivityDate)

PREVIOUS NOTIFICATION:
A first warning was sent to the user and yourself approximately $([math]::Round($midPointThreshold - $userWarningThreshold)) days ago when the account reached $userWarningThreshold days of inactivity.

ACTION REQUIRED:
If this user still needs access to their account:
1. Contact the user directly to ask them to sign in
2. A single successful login will reset the inactivity counter
3. If the user is on extended leave, contact IT to request an exemption

If the user no longer needs this account:
- No action required - the account will be automatically disabled in $daysUntilDisable days
- You will receive a final confirmation notification when the account is disabled

NEXT STEPS:
- Account will be disabled after $inactivityTime days of total inactivity
- You will receive a third notification when/if the account is disabled

NEED HELP?
If you have questions or need to request an exemption, please:
Create a support request via the IT Help Desk

This is an automated notification. Do not reply to this email.
"@
            
            Send-Email -sendGridApiKey $sendGridApiKey `
                -senderEmailAddress $sendGridSenderEmailAddress `
                -recipientEmailAddresses @($user.ManagerEmail) `
                -subject $managerEmailSubject `
                -content $managerEmailContent
            
            Write-Output "Sent second warning (mid-point) to manager: $($user.ManagerEmail) for user: $($user.DisplayName)"
        }
        else {
            Write-Warning "No manager email for user: $($user.DisplayName). Skipping mid-point notification."
        }
    }
    else {
        Write-Output "Testing mode: Would send second notification (mid-point) to manager for user $($user.DisplayName)"
    }
}
