<#
.SYNOPSIS
Updates the membership of a security group for long term leave users based on a CSV file stored in Azure Blob Storage. Sends an email report about successful and failed additions.

.DESCRIPTION
This script reads a CSV file from Azure Blob Storage, removes all members from the specified security group, and adds new members based on their UPN from the CSV. It sends an email report with details of who was added and who failed to add, using SendGrid credentials from Azure Key Vault.

.PARAMETER GroupName
The name of the security group to update.

.PARAMETER StorageAccount
The name of the Azure Storage Account containing the CSV file.

.PARAMETER Container
The name of the blob container.

.PARAMETER BlobName
The name of the CSV file blob.

.PARAMETER SendGridApiKeyKvName
The name of the Key Vault instance containing the SendGrid API key.

.PARAMETER SendGridApiKeyKvSecretName
The name of the secret in the Key Vault containing the SendGrid API key.

.PARAMETER SendGridSenderEmailAddress
The email address of the sender.

.PARAMETER SendGridRecipientEmailAddresses
A comma-separated list of recipient email addresses.

.PARAMETER SendGridApiEndpoint
The endpoint URL for the SendGrid API.
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$GroupName,
    [Parameter(Mandatory=$false)]
    [string]$StorageAccount="storage account name",
    [Parameter(Mandatory=$false)]
    [string]$Container="container name",
    [Parameter(Mandatory=$false)]
    [string]$BlobName="file name.csv",
    [Parameter(Mandatory=$false)]
    [string]$SendGridApiKeyKvName="YourKeyVaultName",
    [Parameter(Mandatory=$false)]
    [string]$SendGridApiKeyKvSecretName="YourSendGridSecretName",
    [Parameter(Mandatory=$false)]
    [string]$SendGridSenderEmailAddress="sender@yourdomain.com",
    [Parameter(Mandatory=$false)]
    [string]$SendGridRecipientEmailAddresses=$null,
    [Parameter(Mandatory=$false)]
    [string]$SendGridApiEndpoint="https://api.sendgrid.com/v3/mail/send"
)

# Automation Account authentication
Connect-AzAccount -Identity
Connect-MgGraph -Identity

# Get blob content using managed identity
$context = New-AzStorageContext -StorageAccountName $StorageAccount -UseManagedIdentity
$csvPath = Join-Path $env:TEMP $BlobName
Get-AzStorageBlobContent -Container $Container -Blob $BlobName -Context $context -Destination $csvPath -Force

# Read and parse CSV
$csvContent = Get-Content $csvPath
$leaveUsers = $csvContent | ConvertFrom-Csv -Delimiter ","

# Get group object
$group = Get-MgGroup -Filter "displayName eq '$GroupName'"
if (-not $group) {
    throw "Group $GroupName not found."
}

# Remove all members from the group
$members = Get-MgGroupMember -GroupId $group.Id
foreach ($member in $members) {
    Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $member.Id
}

# Add members from CSV by UPN
$addedUsers = @()
$unsuccessfulAdds = @()
foreach ($user in $leaveUsers) {
    $upn = ($user.Work_Email).Trim()
    $employeeId = ($user.Employee_ID).Trim()
    $added = $false
    if ([string]::IsNullOrWhiteSpace($upn)) {
        continue
    }
    try {
        $mgUser = Get-MgUser -UserId $upn
        if ($mgUser) {
            New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $mgUser.Id
            $added = $true
            $addedUsers += [PSCustomObject]@{
                Employee_ID = $employeeId
                Work_Email  = $upn
            }
        }
    }
    catch {
        Write-Error "Failed to add user with UPN $upn to group $($group.Id): $($_.Exception.Message)"
    }
    if (-not $added) {
        $unsuccessfulAdds += [PSCustomObject]@{
            Employee_ID = $employeeId
            Work_Email  = $upn
        }
    }
}

# Prepare report files
$addedReportPath = "$env:TEMP\added-users.csv"
$failedReportPath = "$env:TEMP\failed-users.csv"
$addedUsers | Export-Csv -Path $addedReportPath -NoTypeInformation
$unsuccessfulAdds | Export-Csv -Path $failedReportPath -NoTypeInformation

function Send-EmailReport {
    Param(
        [Parameter(Mandatory = $true)]
        [string] $SendGridApiKey,
        [Parameter(Mandatory = $true)]
        [string] $SendGridApiEndpoint,
        [Parameter(Mandatory = $true)]
        [string] $SenderEmailAddress,
        [Parameter(Mandatory = $true)]
        [string[]] $RecipientEmailAddresses,
        [Parameter(Mandatory = $true)]
        [String] $Subject,
        [Parameter(Mandatory = $true)]
        [String] $Content,
        [Parameter(Mandatory = $true)]
        [Object[]] $Attachments
    )

    $headers = @{
        "Authorization" = "Bearer $SendGridApiKey"
        "Content-Type"  = "application/json"
    }

    $attachments = $Attachments | ForEach-Object {
        $contentCsv = Get-Content $_.file -Raw

        @{
            content     = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($contentCsv))
            filename    = [System.IO.Path]::GetFileName($_.file)
            type        = $_.type
            disposition = "attachment"
        }
    }

    $body = @{
        from             = @{ email = $SenderEmailAddress }
        personalizations = @(@{ to = @($RecipientEmailAddresses | ForEach-Object { @{ email = $_ } }) })
        subject          = $Subject
        content          = @(@{ type = "text/plain"; value = $Content })
        attachments      = $attachments
    }

    $bodyJson = $body | ConvertTo-Json -Depth 4
    Invoke-RestMethod -Uri $SendGridApiEndpoint -Method Post -Headers $headers -Body $bodyJson
}

# Get SendGrid API key from Key Vault using managed identity
$sendGridApiKey = (Get-AzKeyVaultSecret -VaultName $SendGridApiKeyKvName -Name $SendGridApiKeyKvSecretName).SecretValueText

# Split comma-separated recipient addresses into an array
$recipientArray = $SendGridRecipientEmailAddresses -split '\s*,\s*'

# Prepare email content
$emailContent = @"
This is a report from update the security group membership for long term leave users.

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

Send-EmailReport -SendGridApiKey $sendGridApiKey `
    -SendGridApiEndpoint $SendGridApiEndpoint `
    -SenderEmailAddress $SendGridSenderEmailAddress `
    -RecipientEmailAddresses $recipientArray `
    -Subject "Group Membership Update Report: $GroupName" `
    -Content $emailContent `
    -Attachments $attachments

