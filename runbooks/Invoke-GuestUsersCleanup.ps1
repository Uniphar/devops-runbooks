<#
.SYNOPSIS
Performs cleanup of inactive guest user accounts in Azure AD based on configurable inactivity thresholds and sends an email report.

.DESCRIPTION
This script identifies Azure AD guest user accounts that have been inactive for a specified number of days.
It can optionally disable or remove these users based on the provided parameters RemoveUsers and DisableUsers.
The script generates CSV reports for all processed users, those to be disabled, those to be removed, and those skipped.
MS Graph Beta API is used to retrieve user details, including last sign-in activity.
If a failure occurs when retrieving users from Microsoft Graph Beta, a warning email is sent using SendGrid.
The script also supports sending email notifications using SendGrid.

.PARAMETER InactiveDaysToRemove
Number of inactive days after which a guest user is considered for removal.

.PARAMETER InactiveDaysToDisableMin
Minimum number of inactive days after which a guest user is considered for disabling.

.PARAMETER RemoveUsers
Boolean. If $true, users identified for removal will be deleted.

.PARAMETER DisableUsers
Boolean. If $true, users identified for disabling will be disabled.

.PARAMETER KeyVaultName
The name of the Azure Key Vault instance containing the SendGrid API key.

.PARAMETER SendGridSecretName
The name of the secret in the Key Vault containing the SendGrid API key.

.PARAMETER ToEmail
The recipient email address for notifications.

.PARAMETER FromEmail
The sender email address for notifications.
#>
[CmdletBinding()]
param (
    [int]$InactiveDaysToRemove = 200,
    [int]$InactiveDaysToDisableMin = 100,
    [bool]$RemoveUsers = $false,
    [bool]$DisableUsers = $false,
    [bool]$TestMode = $true,
    [Parameter(Mandatory = $true)] [string]$KeyVaultName,
    [Parameter(Mandatory = $true)] [string]$SendGridSecretName,
    [Parameter(Mandatory = $true)] [string]$FromEmail,
    [Parameter(Mandatory = $true)] [string]$ToEmail
)

# Connect to Azure using Managed Identity (for Azure Automation)
try {
    Write-Output "Connecting to Azure using Managed Identity..."
    Connect-AzAccount -Identity -ErrorAction Stop
    Write-Output "Successfully connected to Azure using Managed Identity"
} catch {
    Write-Output "Failed to connect to Azure using Managed Identity: $($_.Exception.Message)"
    throw
}

# Get SendGrid API Key from Azure Key Vault if KeyVaultName is provided
if ($KeyVaultName -and $SendGridSecretName) {
    Write-Output "========================================"
    Write-Output "KEY VAULT ACCESS DEBUG INFORMATION:"
    Write-Output "========================================"
    Write-Output "Attempting to retrieve secret from Key Vault..."
    Write-Output "Key Vault Name: $KeyVaultName"
    Write-Output "Secret Name: $SendGridSecretName"
    
    try {
        # First, try to check if we can access the Key Vault
        Write-Output "Testing Key Vault access..."
        $kvTest = Get-AzKeyVault -VaultName $KeyVaultName -ErrorAction Stop
        Write-Output "✓ Key Vault found: $($kvTest.VaultName) (Resource Group: $($kvTest.ResourceGroupName))"
        
        # Now try to get the secret
        Write-Output "Attempting to retrieve secret '$SendGridSecretName'..."
        $secret = Get-AzKeyVaultSecret -VaultName $KeyVaultName -Name $SendGridSecretName -AsPlainText -ErrorAction Stop
        $SendGridApiKey = $secret
        Write-Output "✓ SendGrid API Key retrieved successfully from Key Vault '$KeyVaultName'"
        Write-Output "✓ Secret length: $($SendGridApiKey.Length) characters"
    } catch {
        Write-Output "✗ Failed to retrieve SendGrid API Key from Key Vault '$KeyVaultName'"
        Write-Output "Error Type: $($_.Exception.GetType().FullName)"
        Write-Output "Error Message: $($_.Exception.Message)"
        
        if ($_.Exception.Message -like "*403*" -or $_.Exception.Message -like "*Forbidden*" -or $_.Exception.Message -like "*not authorized*") {
            Write-Output ""
            Write-Output "PERMISSION ISSUE DETECTED:"
            Write-Output "The Managed Identity does not have permission to access Key Vault secrets."
            Write-Output ""
            Write-Output "TO FIX THIS:"
            Write-Output "1. Go to Azure Portal > Key Vault: $KeyVaultName"
            Write-Output "2. Navigate to 'Access policies' or 'Access control (IAM)'"
            Write-Output "3. Add the Automation Account's Managed Identity with 'Get' permission for Secrets"
            Write-Output "   OR assign 'Key Vault Secrets User' role to the Managed Identity"
        } elseif ($_.Exception.Message -like "*not found*" -or $_.Exception.Message -like "*404*") {
            Write-Output ""
            Write-Output "SECRET NOT FOUND:"
            Write-Output "The secret '$SendGridSecretName' does not exist in Key Vault '$KeyVaultName'"
            Write-Output "Please verify the secret name is correct."
        }
        
        $SendGridApiKey = ""
    }
    Write-Output "========================================"
} else {
    Write-Output "Key Vault Name or Secret Name not provided. Skipping Key Vault access."
    $SendGridApiKey = ""
}

$scriptStart = Get-Date

# Connect to Microsoft Graph using Managed Identity (for Azure Automation)
try {
    Write-Output "Connecting to Microsoft Graph using Managed Identity..."
    Connect-MgGraph -Identity -ErrorAction Stop
    Write-Output "Successfully connected to Microsoft Graph using Managed Identity"
} catch {
    Write-Output "Failed to connect to Microsoft Graph using Managed Identity: $($_.Exception.Message)"
    throw
}

$InactiveDaysToDisableMax = $InactiveDaysToRemove

# Get all guest users
Write-Output "========================================"
Write-Output "RETRIEVING GUEST USERS FROM MICROSOFT GRAPH:"
Write-Output "========================================"
Write-Output "Attempting to query guest users with SignInActivity..."
Write-Output "Filter: userType eq 'Guest'"
Write-Output "Properties: Id, UserPrincipalName, CreatedDateTime, LastPasswordChangeDateTime, SignInActivity"
Write-Output ""

try {
    $guests = Get-MgBetaUser -Filter "userType eq 'Guest'" `
        -Property "Id,UserPrincipalName,CreatedDateTime,LastPasswordChangeDateTime,SignInActivity" `
        -All -ErrorAction Stop
    
    Write-Output "✓ Successfully retrieved $($guests.Count) guest users."
    
    Write-Output "========================================"
} catch {
    $errorMsg = "ERROR: Failed to retrieve users from Microsoft Graph Beta API. Exception: $($_.Exception.Message)"
    Write-Output $errorMsg
    Write-Output ""
    Write-Output "Error Type: $($_.Exception.GetType().FullName)"
    Write-Output "Error Category: $($_.CategoryInfo.Category)"
    Write-Output ""
    
    # Specific troubleshooting based on error
    if ($_.Exception.Message -like "*AuditLog.Read.All*") {
        Write-Output "ISSUE IDENTIFIED: AuditLog.Read.All permission error"
        Write-Output ""
        Write-Output "The error indicates that the Managed Identity for the Automation Account is missing the 'AuditLog.Read.All' MS Graph API permission."
        Write-Output "A Global Admin must grant this permission in Azure AD / Enterprise Applications."
        Write-Output ""
    }
    
    Write-Output "========================================"

    # Prepare SendGrid email payload
    $emailBody = @{
        personalizations = @(@{ to = @(@{ email = $ToEmail }) })
        from             = @{ email = $FromEmail }
        subject          = "Inactive Guest Accounts Cleanup - MS Graph Beta Request Failed"
        content          = @(@{ type = "text/plain"; value = $errorMsg })
    } | ConvertTo-Json -Depth 4

    if ($SendGridApiKey) {
        try {
            Invoke-RestMethod -Uri "https://api.sendgrid.com/v3/mail/send" `
                -Method Post `
                -Headers @{ "Authorization" = "Bearer $SendGridApiKey"; "Content-Type" = "application/json" } `
                -Body $emailBody
            Write-Output "Warning email sent to $ToEmail."
        } catch {
            Write-Output "Failed to send warning email via SendGrid: $($_.Exception.Message)"
        }
    } else {
        Write-Output "SendGrid API Key not available. Cannot send warning email."
    }

    # Stop the script
    throw "Stopping script due to MS Graph Beta request failure."
}

$now = Get-Date
$toRemove = @()
$toDisable = @()
$skippedUsers = @()
$report = @()
$actionLog = @()

# Process each guest user
foreach ($guest in $guests) {
    $lastSignIn = $null
    $reportLastSignIn = ""
    $inactiveDays = $null

    # Use SignInActivity from Microsoft Graph Beta
    if ($guest.SignInActivity -and $guest.SignInActivity.LastSignInDateTime) {
        $lastSignIn = $guest.SignInActivity.LastSignInDateTime
        $reportLastSignIn = $lastSignIn
        Write-Output "[$($guest.UserPrincipalName)] LastSignInDateTime from Graph: $lastSignIn"
    } elseif ($guest.CreatedDateTime) {
        $lastSignIn = $guest.CreatedDateTime
        $reportLastSignIn = "no sign in ever"
        Write-Output "[$($guest.UserPrincipalName)] Using CreatedDateTime for inactivity calculation: $lastSignIn"
    } elseif ($guest.LastPasswordChangeDateTime) {
        $lastSignIn = $guest.LastPasswordChangeDateTime
        $reportLastSignIn = "no sign in ever"
        Write-Output "[$($guest.UserPrincipalName)] Using LastPasswordChangeDateTime for inactivity calculation: $lastSignIn"
    } else {
        Write-Output "[$($guest.UserPrincipalName)] No usable date found for inactivity calculation. Skipping user."
        $skippedUsers += [PSCustomObject]@{
            UserPrincipalName = $guest.UserPrincipalName
            Id                = $guest.Id
            Reason            = "No sign-in activity, creation date, or password change date"
        }
        continue
    }

    # Ensure $lastSignIn is a DateTime object
    if ($lastSignIn -is [System.DateTimeOffset]) {
        $lastSignInValue = $lastSignIn.DateTime
    } else {
        $lastSignInValue = $lastSignIn
    }

    if (-not ($lastSignInValue -is [DateTime])) {
        try {
            $lastSignInValue = [datetime]$lastSignInValue
            Write-Output "[$($guest.UserPrincipalName)] Final used date for inactivity calculation: $lastSignInValue"
        } catch {
            Write-Output "[$($guest.UserPrincipalName)] Could not convert date for inactivity calculation. Skipping user."
            continue
        }
    } else {
        Write-Output "[$($guest.UserPrincipalName)] Final used date for inactivity calculation: $lastSignInValue"
    }

    $inactiveDays = ($now - $lastSignInValue).Days

    # Build a report object for this guest
    $report += [PSCustomObject]@{
        UserPrincipalName = $guest.UserPrincipalName
        Id                = $guest.Id
        LastSignInDate    = $reportLastSignIn
        CreationDate      = $guest.CreatedDateTime
        InactiveDays      = $inactiveDays
    }

    # Classify users based on inactivity days.
    if ($inactiveDays -gt $InactiveDaysToRemove) {
        $toRemove += $guest
    } elseif (($inactiveDays -ge $InactiveDaysToDisableMin) -and ($inactiveDays -le $InactiveDaysToDisableMax)) {
        $toDisable += $guest
    }
}

# Display summary statistics
Write-Output "========================================"
Write-Output "GUEST USER ANALYSIS SUMMARY:"
Write-Output "========================================"
Write-Output "Total guest users found: $($guests.Count)"
Write-Output "Users to be disabled (inactive $InactiveDaysToDisableMin-$InactiveDaysToDisableMax days): $($toDisable.Count)"
Write-Output "Users to be removed (inactive >$InactiveDaysToRemove days): $($toRemove.Count)"
Write-Output "Skipped users: $($skippedUsers.Count)"
Write-Output ""
Write-Output "Inactivity breakdown:"
$inactivityGroups = $report | Group-Object { 
    if ($_.InactiveDays -le 30) { "0-30 days" }
    elseif ($_.InactiveDays -le 60) { "31-60 days" }
    elseif ($_.InactiveDays -le 90) { "61-90 days" }
    elseif ($_.InactiveDays -le 120) { "91-120 days" }
    elseif ($_.InactiveDays -le 180) { "121-180 days" }
    elseif ($_.InactiveDays -le 365) { "181-365 days" }
    else { "365+ days" }
}
foreach ($group in ($inactivityGroups | Sort-Object Name)) {
    Write-Output "  $($group.Name): $($group.Count) users"
}
Write-Output "========================================"

# Set export directory to temp path suitable for Azure Automation
$exportDir = $env:TEMP
if (-not (Test-Path $exportDir)) {
    New-Item -Path $exportDir -ItemType Directory | Out-Null
}
Write-Output "Export directory: $exportDir"

# Export all guest user objects with only available details for debugging
$guests | Select-Object UserPrincipalName, CreatedDateTime, LastPasswordChangeDateTime |
    ForEach-Object {
        $g = $_
        $reportEntry = $report | Where-Object { $_.UserPrincipalName -eq $g.UserPrincipalName }
        [PSCustomObject]@{
            UserPrincipalName          = $g.UserPrincipalName
            CreatedDateTime            = $g.CreatedDateTime
            LastPasswordChangeDateTime = $g.LastPasswordChangeDateTime
            LastSignInDate             = if ($reportEntry) { $reportEntry.LastSignInDate } else { $null }
            InactiveDays               = if ($reportEntry) { $reportEntry.InactiveDays } else { $null }
        }
    } | Export-Csv -Path "$exportDir\AllGuestUsersDebug.csv" -NoTypeInformation
Write-Output "All guest user details exported for debugging: $exportDir\AllGuestUsersDebug.csv"

# Export users to be disabled
$toDisable | ForEach-Object {
    $guest = $_
    $reportEntry = $report | Where-Object { $_.UserPrincipalName -eq $guest.UserPrincipalName }
    [PSCustomObject]@{
        UserPrincipalName          = $guest.UserPrincipalName
        CreatedDateTime            = $guest.CreatedDateTime
        LastPasswordChangeDateTime = $guest.LastPasswordChangeDateTime
        LastSignInDate             = if ($reportEntry) { $reportEntry.LastSignInDate } else { $null }
        InactiveDays               = if ($reportEntry) { $reportEntry.InactiveDays } else { $null }
    }
} | Export-Csv -Path "$exportDir\GuestsToDisable.csv" -NoTypeInformation
Write-Output "Users to be disabled exported: $exportDir\GuestsToDisable.csv"

# Export users to be removed
$toRemove | ForEach-Object {
    $guest = $_
    $reportEntry = $report | Where-Object { $_.UserPrincipalName -eq $guest.UserPrincipalName }
    [PSCustomObject]@{
        UserPrincipalName          = $guest.UserPrincipalName
        CreatedDateTime            = $guest.CreatedDateTime
        LastPasswordChangeDateTime = $guest.LastPasswordChangeDateTime
        LastSignInDate             = if ($reportEntry) { $reportEntry.LastSignInDate } else { $null }
        InactiveDays               = if ($reportEntry) { $reportEntry.InactiveDays } else { $null }
    }
} | Export-Csv -Path "$exportDir\GuestsToRemove.csv" -NoTypeInformation
Write-Output "Users to be removed exported: $exportDir\GuestsToRemove.csv"

# Perform disable and remove actions only if parameters are true
if ($DisableUsers -and $toDisable.Count -gt 0) {
    foreach ($user in $toDisable) {
        if ($TestMode) {
            Write-Output "[TESTMODE] Would disable user: $($user.UserPrincipalName)"
            $actionLog += "[TESTMODE] Would disable user: $($user.UserPrincipalName) (Id: $($user.Id)) at $(Get-Date -Format o)"
        } else {
            Write-Output "Disabling user: $($user.UserPrincipalName)"
            try {
                Update-MgBetaUser -UserId $user.Id -AccountEnabled:$false
                $actionLog += "Disabled user: $($user.UserPrincipalName) (Id: $($user.Id)) at $(Get-Date -Format o)"
            } catch {
                $actionLog += "Failed to disable user: $($user.UserPrincipalName) (Id: $($user.Id)) at $(Get-Date -Format o) - Error: $($_.Exception.Message)"
            }
        }
    }
} else {
    foreach ($user in $toDisable) {
        $actionLog += "User NOT disabled (DisableUsers is false): $($user.UserPrincipalName) (Id: $($user.Id)) at $(Get-Date -Format o)"
    }
    if ($toDisable.Count -eq 0) {
        $actionLog += "No users to disable at $(Get-Date -Format o)"
    }
    Write-Output "User disabling is skipped (DisableUsers is false or no users to disable)."
}

if ($RemoveUsers -and $toRemove.Count -gt 0) {
    foreach ($user in $toRemove) {
        if ($TestMode) {
            Write-Output "[TESTMODE] Would remove user: $($user.UserPrincipalName)"
            $actionLog += "[TESTMODE] Would remove user: $($user.UserPrincipalName) (Id: $($user.Id)) at $(Get-Date -Format o)"
        } else {
            Write-Output "Removing user: $($user.UserPrincipalName)"
            try {
                Remove-MgBetaUser -UserId $user.Id -Confirm:$false
                $actionLog += "Removed user: $($user.UserPrincipalName) (Id: $($user.Id)) at $(Get-Date -Format o)"
            } catch {
                $actionLog += "Failed to remove user: $($user.UserPrincipalName) (Id: $($user.Id)) at $(Get-Date -Format o) - Error: $($_.Exception.Message)"
            }
        }
    }
} else {
    foreach ($user in $toRemove) {
        $actionLog += "User NOT removed (RemoveUsers is false): $($user.UserPrincipalName) (Id: $($user.Id)) at $(Get-Date -Format o)"
    }
    if ($toRemove.Count -eq 0) {
        $actionLog += "No users to remove at $(Get-Date -Format o)"
    }
    Write-Output "User removal is skipped (RemoveUsers is false or no users to remove)."
}

# Write action log to file
$logPath = "$exportDir\ActionLog.txt"
$actionLog | Out-File -FilePath $logPath -Encoding utf8

# Export skipped users if any
if ($skippedUsers.Count -gt 0) {
    $skippedUsers | Export-Csv -Path "$exportDir\SkippedGuests.csv" -NoTypeInformation
    Write-Output "Skipped users report exported: $exportDir\SkippedGuests.csv"
}

$scriptEnd = Get-Date
$duration = $scriptEnd - $scriptStart
Write-Output "Script runtime: $($duration.TotalSeconds) seconds"

# Send all reports and log as attachments in one email via SendGrid
# Build attachments list defensively (only include files that exist)
$attachments = @()
$filesToAttach = @("$exportDir\AllGuestUsersDebug.csv", "$exportDir\GuestsToDisable.csv", "$exportDir\GuestsToRemove.csv", $logPath)
foreach ($f in $filesToAttach) {
    if (Test-Path $f) {
        $attachments += @{ content = [Convert]::ToBase64String([IO.File]::ReadAllBytes($f)); filename = [IO.Path]::GetFileName($f); type = "text/csv"; disposition = "attachment" }
    } else {
        Write-Output "Attachment missing, skipping: $f"
    }
}

$emailSubject = if ($TestMode) {
    "Inactive Guest Accounts Cleanup - Report [TEST MODE]"
} else {
    "Inactive Guest Accounts Cleanup - Report"
}

$emailContent = if ($TestMode) {
    "TEST MODE: No users were actually disabled or removed. See attached reports showing what WOULD happen in production mode."
} else {
    "See attached reports and action log."
}

# Debug: Email configuration
Write-Output "========================================"
Write-Output "EMAIL SENDING DEBUG INFORMATION:"
Write-Output "========================================"
Write-Output "SendGrid API Key present: $(if ($SendGridApiKey) { 'YES (length: ' + $SendGridApiKey.Length + ')' } else { 'NO' })"
Write-Output "From Email: $FromEmail"
Write-Output "To Email: $ToEmail"
Write-Output "Subject: $emailSubject"
Write-Output "Number of attachments: $($attachments.Count)"
Write-Output "Attachment files:"
foreach ($att in $attachments) {
    Write-Output "  - $($att.filename) (size: $($att.content.Length) base64 chars)"
}
Write-Output "Test Mode: $TestMode"
Write-Output "========================================"

$emailBody = @{
    personalizations = @(@{ to = @(@{ email = $ToEmail }) })
    from             = @{ email = $FromEmail }
    subject          = $emailSubject
    content          = @(@{ type = "text/plain"; value = $emailContent })
    attachments      = $attachments
} | ConvertTo-Json -Depth 6

# Debug: Show email body size
Write-Output "Email body JSON size: $($emailBody.Length) characters"

# Send email in both test mode and production mode
if ($SendGridApiKey) {
    Write-Output "Attempting to send email via SendGrid..."
    try {
        $response = Invoke-RestMethod -Uri "https://api.sendgrid.com/v3/mail/send" `
            -Method Post `
            -Headers @{ "Authorization" = "Bearer $SendGridApiKey"; "Content-Type" = "application/json" } `
            -Body $emailBody
        
        Write-Output "SendGrid API Response: $($response | ConvertTo-Json -Compress)"
        
        if ($TestMode) {
            Write-Output "[TESTMODE] ✓ Summary email with reports and action log sent successfully to $ToEmail."
        } else {
            Write-Output "✓ Summary email with reports and action log sent successfully to $ToEmail."
        }
    } catch {
        Write-Output "✗ Failed to send summary email via SendGrid"
        Write-Output "Error Message: $($_.Exception.Message)"
        Write-Output "Error Details: $($_.Exception.Response.StatusCode) - $($_.Exception.Response.StatusDescription)"
        if ($_.Exception.Response) {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd()
            Write-Output "SendGrid Response Body: $responseBody"
        }
    }
} else {
    Write-Output "✗ SendGrid API Key not available. Cannot send summary email."
    Write-Output "Please check KeyVault configuration: KeyVaultName='$KeyVaultName', SecretName='$SendGridSecretName'"
}
Write-Output "========================================"
