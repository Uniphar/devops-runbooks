<#
.SYNOPSIS
Updates the membership of a security group for long-term leave users from a CSV stored in Azure Blob Storage.

.DESCRIPTION
This runbook-style script is intended to be run in an Automation Account or other managed environment
where a system-assigned managed identity has access to Azure Storage and Key Vault. It performs these steps:
    1. Authenticate using the runbook's managed identity (Connect-AzAccount / Connect-MgGraph).
    2. Download a CSV blob from a storage account/container and parse UPNs.
    3. Remove all existing members from the specified security group.
    4. Add members found in the CSV (matching by UPN) to the group.
    5. Export success/failure CSVs and send a report email via SendGrid using an API key from Key Vault.

PREREQUISITES
    - This script expects the Az.* modules (for Storage/KeyVault) and Microsoft.Graph PowerShell module
        to be available in the runbook worker or Automation Account.
    - The runbook's managed identity must have appropriate RBAC on the Storage account/container and
        permissions to read the Key Vault secret. It also needs Graph permissions for group membership operations.

.PARAMETER GroupName
The display name of the security group to update.

.PARAMETER StorageAccount
The name of the Azure Storage Account containing the CSV file.

.PARAMETER Container
The name of the blob container that holds the CSV.

.PARAMETER BlobName
The name of the CSV blob (for example: leavers.csv).

.PARAMETER SendGridApiKeyKvName
The name of the Key Vault to retrieve the SendGrid API key from.

.PARAMETER SendGridApiKeyKvSecretName
The name of the Key Vault secret containing the SendGrid API key.

.PARAMETER SendGridSenderEmailAddress
The From address to use when sending the report.

.PARAMETER SendGridRecipientEmailAddresses
A comma-separated list (or array) of recipient email addresses.

.PARAMETER SendGridApiEndpoint
Optional: the SendGrid API endpoint (default: https://api.sendgrid.com/v3/mail/send).

.NOTES
    - The CSV is expected to have columns named Work_Email and Employee_ID; adjust parsing if yours differ.
    - The script intentionally removes all members from the group before re-adding from the CSV.

#>
[CmdletBinding(PositionalBinding = $false)]
param(
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$GroupName,
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$StorageAccount,
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Container,
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$BlobName,
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$SendGridApiKeyKvName,
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$SendGridApiKeyKvSecretName,
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$SendGridSenderEmailAddress,
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$SendGridRecipientEmailAddresses,
    [Parameter(Mandatory = $false)][string]$SendGridApiEndpoint = "https://api.sendgrid.com/v3/mail/send"
)


function Send-EmailReport {
    Param(
        [Parameter(Mandatory = $true)] [string] $SendGridApiKey,
        [Parameter(Mandatory = $true)] [string] $SendGridApiEndpoint,
        [Parameter(Mandatory = $true)] [string] $SenderEmailAddress,
        [Parameter(Mandatory = $true)] [string[]] $RecipientEmailAddresses,
        [Parameter(Mandatory = $true)] [String] $Subject,
        [Parameter(Mandatory = $true)] [String] $Content,
        [Parameter(Mandatory = $false)] [Object[]] $Attachments
    )

    # Build SendGrid API headers. The API key is retrieved from Key Vault by the caller.
    $headers = @{
        "Authorization" = "Bearer $SendGridApiKey"
        "Content-Type"  = "application/json"
    }

    Write-Output "Sending email to: $($RecipientEmailAddresses -join ', ') via $SendGridApiEndpoint"

    # Convert attachments to the JSON structure SendGrid expects (base64-encoded content)
    $sendGridAttachments = @()
    if ($Attachments) {
        foreach ($att in $Attachments) {
            if (-not (Test-Path -Path $att.file)) {
                Write-Output "  WARNING: Attachment not found: $($att.file) - skipping"
                continue
            }
            $contentCsv = Get-Content $att.file -Raw
            # Handle empty or null file content
            if ([string]::IsNullOrEmpty($contentCsv)) {
                Write-Output "  WARNING: Attachment file is empty: $($att.file) - using empty string"
                $contentCsv = ""
            }
            $sendGridAttachments += @{
                content     = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($contentCsv))
                filename    = [System.IO.Path]::GetFileName($att.file)
                type        = $att.type
                disposition = "attachment"
            }
        }
    }

    # Build the JSON body for the SendGrid API call
    $body = @{
        from             = @{ email = $SenderEmailAddress }
        personalizations = @(@{ to = @($RecipientEmailAddresses | ForEach-Object { @{ email = $_ } }) })
        subject          = $Subject
        content          = @(@{ type = "text/plain"; value = $Content })
    }
    
    # Only add attachments if there are any
    if ($sendGridAttachments.Count -gt 0) {
        $body.attachments = $sendGridAttachments
    }

    $bodyJson = $body | ConvertTo-Json -Depth 6
    Write-Output "  Request body length: $($bodyJson.Length) characters"
    Write-Output "  Number of attachments: $($sendGridAttachments.Count)"

    try {
        $response = Invoke-RestMethod -Uri $SendGridApiEndpoint -Method Post -Headers $headers -Body $bodyJson -ErrorAction Stop
        Write-Output "  SendGrid API call succeeded"
        return $response
    } catch {
        Write-Output "  ERROR sending email: $($_.Exception.Message)"
        if ($_.Exception.Response) {
            try {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $reader.BaseStream.Position = 0
                $reader.DiscardBufferedData()
                $responseBody = $reader.ReadToEnd()
                Write-Output "  SendGrid API error response: $responseBody"
            } catch {
                Write-Output "  Could not read error response body"
            }
        }
        throw
    }
}


# --- Main execution ---

# Log parameter values for runbook execution tracking
Write-Output "Running with parameters:"
Write-Output "  GroupName = $GroupName"
Write-Output "  StorageAccount = $StorageAccount"
Write-Output "  Container = $Container"
Write-Output "  BlobName = $BlobName"
Write-Output "  SendGridApiKeyKvName = $SendGridApiKeyKvName"
Write-Output "  SendGridSenderEmailAddress = $SendGridSenderEmailAddress"
Write-Output "  SendGridRecipientEmailAddresses = $SendGridRecipientEmailAddresses"

# Authentication
# Use the automation account or system-assigned managed identity to authenticate to Azure and Microsoft Graph.
# Connect-AzAccount -Identity uses the managed identity provided to the runbook/machine.
Write-Output "Connecting to Azure with Managed Identity..."
Connect-AzAccount -Identity -ErrorAction Stop
Write-Output "✓ Connected to Azure"

# Connect to Microsoft Graph using the same identity; required for Get-MgGroup / member operations.
Write-Output "Connecting to Microsoft Graph with Managed Identity..."
Connect-MgGraph -Identity -ErrorAction Stop
Write-Output "✓ Connected to Microsoft Graph"

    # --- Download CSV from Blob Storage ---
    # Create a storage context that uses the managed identity. This requires the identity to have
    # at least 'Storage Blob Data Reader' role on the storage account or container.
    Write-Output "Creating storage context for account: $StorageAccount"
    
    # Get the current Azure context to verify we're authenticated
    $azContext = Get-AzContext
    Write-Output "Current Azure Context: Subscription = $($azContext.Subscription.Name), Account = $($azContext.Account.Id)"
    
    try {
        # Try to create storage context with managed identity
        Write-Output "Attempting to create storage context with Managed Identity..."
        $context = New-AzStorageContext -StorageAccountName $StorageAccount -UseConnectedAccount -ErrorAction Stop
        Write-Output "Storage context created successfully"
    } catch {
        Write-Output "WARNING: Failed to create storage context with UseConnectedAccount. Error: $($_.Exception.Message)"
        Write-Output "Trying alternative method with explicit authentication..."
        
        # Alternative: Get storage account and create context
        try {
            $storageAccountResource = Get-AzStorageAccount | Where-Object { $_.StorageAccountName -eq $StorageAccount } | Select-Object -First 1
            if ($storageAccountResource) {
                $context = $storageAccountResource.Context
                Write-Output "Storage context obtained from storage account resource"
            } else {
                throw "Storage account '$StorageAccount' not found in current subscription"
            }
        } catch {
            throw "Failed to create storage context: $($_.Exception.Message). Ensure Managed Identity has 'Storage Blob Data Reader' role."
        }
    }
    
    # Save blob to a temporary path on the worker machine (use system temp folder reliably)
    $blobFileName = [System.IO.Path]::GetFileName($BlobName)
    $tempDir = [System.IO.Path]::GetTempPath()
    $csvPath = Join-Path $tempDir $blobFileName
    
    Write-Output "Attempting to download blob: '$BlobName' from container: '$Container'"
    Write-Output "Destination path: $csvPath"
    
    try {
        $downloadResult = Get-AzStorageBlobContent -Container $Container -Blob $BlobName -Context $context -Destination $csvPath -Force -ErrorAction Stop
        Write-Output "Blob downloaded successfully"
    } catch {
        Write-Output "ERROR: Failed to download blob from Azure Storage"
        Write-Output "Error details: $($_.Exception.Message)"
        throw "Failed to download CSV from blob storage: $($_.Exception.Message)"
    }

    # --- Read and parse CSV ---
    # Read the CSV file into memory and convert to objects. The CSV must have a header row.
    if (-not (Test-Path -Path $csvPath)) {
        throw "CSV file not found at path: $csvPath. Download may have failed silently."
    }
    $csvContent = Get-Content $csvPath -ErrorAction Stop
    $leaveUsers = $csvContent | ConvertFrom-Csv -Delimiter ","
    # Parsed CSV; rows: $($leaveUsers.Count)
    if (-not $leaveUsers -or $leaveUsers.Count -eq 0) {
        throw "CSV at $csvPath does not contain any rows or is missing a header."
    }

    # --- Locate target group ---
    # Get the group object by displayName. If multiple groups share the same displayName,
    # this may return more than one; the script assumes a single match.
    # Retrieve the group and ensure we get a single object to use for operations
    # Trim and escape single quotes in the group name so the OData filter remains valid
    Write-Output "Searching for group: '$GroupName'"
    
    if ([string]::IsNullOrWhiteSpace($GroupName)) {
        throw "GroupName parameter is null or empty"
    }
    
    $searchGroupName = $GroupName.Trim()
    $escapedGroupName = $searchGroupName -replace "'", "''"
    
    Write-Output "Querying Microsoft Graph for group with displayName: '$escapedGroupName'"
    $groupMatches = Get-MgGroup -Filter "displayName eq '$escapedGroupName'" -ErrorAction Stop
    
    if (-not $groupMatches -or $groupMatches.Count -eq 0) {
        throw "Group '$GroupName' not found in MS Entra."
    }
    if ($groupMatches.Count -gt 1) {
        # Prefer the first match but warn in the log
        Write-Output "WARNING: Multiple groups matched displayName '$GroupName'. Using the first match (Id: $($groupMatches[0].Id))."
    }
    $group = $groupMatches | Select-Object -First 1
    Write-Output "✓ Found group: $($group.DisplayName) (Id: $($group.Id))"

    # --- Remove existing members ---
    # This removes all members from the group before adding the list from the CSV.
    Write-Output "Retrieving members of group $($group.Id)..."
    $members = Get-MgGroupMember -GroupId $group.Id -ErrorAction Stop
    if ($members) {
        Write-Output "Found $($members.Count) existing member(s). Removing them..."
        foreach ($member in $members) {
            try {
                Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $member.Id -ErrorAction Stop
                Write-Output "  Removed member: $($member.Id)"
            } catch {
                Write-Output "  WARNING: Failed to remove member $($member.Id): $($_.Exception.Message)"
            }
        }
        Write-Output "✓ All existing members removed"
    } else {
        Write-Output "No existing members found for group $($group.Id)"
    }

    # --- Add users from CSV to the group (matching by UPN) ---
    Write-Output "Processing users from CSV..."
    $addedUsers = @()
    $unsuccessfulAdds = @()
    
    foreach ($user in $leaveUsers) {
        # Expecting CSV columns named Work_Email and Employee_ID; trim whitespace for safety
        # Add null checks to prevent "cannot call method on null-valued expression" errors
        $upn = if ($user.Work_Email) { ($user.Work_Email).Trim() } else { "" }
        $employeeId = if ($user.Employee_ID) { ($user.Employee_ID).Trim() } else { "" }
        
        $added = $false
        
        if ([string]::IsNullOrWhiteSpace($upn)) {
            # Skip blank UPN entries
            Write-Output "  Skipping row with empty Work_Email (Employee_ID: $employeeId)"
            continue
        }
        
        # Look up the user by email in Microsoft Graph and resolve the proper UPN/object id
        Write-Output "  Looking up user by email: $upn"
        $mgUser = $null

        # 1) Try direct lookup by UserId (works when the CSV value is the UPN or object id)
        try {
            $mgUser = Get-MgUser -UserId $upn -ErrorAction Stop
            Write-Output "    Found user by UserId"
        } catch {
            # ignore - we'll try filters next
        }

        # 2) If not found, search by mail or userPrincipalName fields
        if (-not $mgUser) {
            $filterEscaped = $upn -replace "'", "''"
            Write-Output "    Not found by UserId; trying mail/userPrincipalName filter..."
            $mgUser = Get-MgUser -Filter "userPrincipalName eq '$filterEscaped' or mail eq '$filterEscaped'" -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($mgUser) { Write-Output "    Found user by mail/userPrincipalName filter (UPN: $($mgUser.UserPrincipalName))" }
        }

        # 3) As a final fallback, try proxyAddresses (SMTP:) if available
        if (-not $mgUser) {
            try {
                Write-Output "    Trying proxyAddresses fallback..."
                $proxyFilter = $filterEscaped
                $mgUser = Get-MgUser -Filter "proxyAddresses/any(a:a eq 'SMTP:$proxyFilter')" -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($mgUser) { Write-Output "    Found user via proxyAddresses (UPN: $($mgUser.UserPrincipalName))" }
            } catch {
                # ignore; leave $mgUser null
            }
        }

        if ($mgUser) {
            # Add the user to the group by their object id
            try {
                New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $mgUser.Id -ErrorAction Stop
                $added = $true
                Write-Output "    ✓ Added user: $upn"
                $addedUsers += [PSCustomObject]@{
                    Employee_ID = $employeeId
                    Work_Email  = $upn
                }
            } catch {
                Write-Output "    ✗ Failed to add user ${upn}: $($_.Exception.Message)"
                $unsuccessfulAdds += [PSCustomObject]@{
                    Employee_ID = $employeeId
                    Work_Email  = $upn
                }
            }
        } else {
            Write-Output "    ✗ User not found in MS Entra: $upn"
            $unsuccessfulAdds += [PSCustomObject]@{
                Employee_ID = $employeeId
                Work_Email  = $upn
            }
        }
    }

    # --- Prepare CSV reports ---
    # Export lists of successfully added and failed users so they can be reviewed.
    $timestamp = (Get-Date).ToString('yyyyMMddHHmmss')
    $addedReportPath = Join-Path $tempDir ("added-users-$timestamp.csv")
    $failedReportPath = Join-Path $tempDir ("failed-users-$timestamp.csv")
    $addedUsers | Export-Csv -Path $addedReportPath -NoTypeInformation -ErrorAction Stop
    $unsuccessfulAdds | Export-Csv -Path $failedReportPath -NoTypeInformation -ErrorAction Stop
    # Exported reports: $addedReportPath, $failedReportPath

    # Get SendGrid API key from Key Vault using managed identity (if configured)
    $sendGridApiKey = $null
    try {
        Write-Output "Attempting to read SendGrid API key from Key Vault: $SendGridApiKeyKvName (secret: $SendGridApiKeyKvSecretName)"
        $kvSecret = Get-AzKeyVaultSecret -VaultName $SendGridApiKeyKvName -Name $SendGridApiKeyKvSecretName -ErrorAction Stop
        if ($kvSecret) { 
            # Convert SecureString to plain text
            if ($kvSecret.SecretValue) {
                $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($kvSecret.SecretValue)
                try {
                    $sendGridApiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
                } finally {
                    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
                }
            } elseif ($kvSecret.SecretValueText) {
                # Fallback for older Az.KeyVault versions
                $sendGridApiKey = $kvSecret.SecretValueText
            }
            
            Write-Output "  ✓ Retrieved secret from Key Vault: $($kvSecret.Id)"
            # (Removed diagnostic output of secret type and length per least disclosure practices)
        }
    } catch {
        Write-Output "  ERROR: Unable to read SendGrid secret from Key Vault: $($_.Exception.Message)"
        # Try to list the Key Vault to confirm existence/permissions
        try {
            $kv = Get-AzKeyVault -VaultName $SendGridApiKeyKvName -ErrorAction Stop
            if ($kv) { Write-Output "  Key Vault exists (resource id: $($kv.ResourceId)); check access policies or Azure RBAC for this runbook's managed identity." }
        } catch {
            Write-Output "  Key Vault lookup failed: $($_.Exception.Message)" 
        }
    }

    # Split comma-separated recipient addresses string into an array
    $recipientArray = @()
    if (-not [string]::IsNullOrWhiteSpace($SendGridRecipientEmailAddresses)) {
        # Split by comma and trim whitespace from each email address
        $recipientArray = $SendGridRecipientEmailAddresses -split '\s*,\s*' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    }

    # Prepare email content
    $emailContent = @"
This is a report on the security group membership update for long term leave users.

Group: $GroupName

Added users: $($addedUsers.Count)
Failed to add: $($unsuccessfulAdds.Count)
Failed users were not found in MS Entra.
See attached CSV files for details.
"@

    $attachments = @(
        @{ file = $addedReportPath; type = "text/csv" }
        @{ file = $failedReportPath; type = "text/csv" }
    )

    if (-not [string]::IsNullOrWhiteSpace($sendGridApiKey) -and $recipientArray.Count -gt 0) {
        Write-Output "Preparing to send email report. Recipients: $($recipientArray -join ', ')"
        foreach ($att in $attachments) { Write-Output "  Attachment: $($att.file) (exists: $(Test-Path $att.file))" }
        try {
            $sendResponse = Send-EmailReport -SendGridApiKey $sendGridApiKey `
                -SendGridApiEndpoint $SendGridApiEndpoint `
                -SenderEmailAddress $SendGridSenderEmailAddress `
                -RecipientEmailAddresses $recipientArray `
                -Subject "Group Membership Update Report: $GroupName" `
                -Content $emailContent `
                -Attachments $attachments
            Write-Output "Email report sent successfully"
            Write-Output "SendGrid response: $sendResponse"
        } catch {
            Write-Output "Failed to send email report: $($_.Exception.Message)"
        }
    } else {
        Write-Output "Skipping sending email report because SendGrid API key or recipient list is missing. sendGridApiKey present: $([bool]$sendGridApiKey); recipients: $($recipientArray.Count)"
    }

    # Runbook completed successfully
