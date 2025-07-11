<#
.SYNOPSIS
Performs cleanup of devices based on inactivity thresholds and sends an email report.

.DESCRIPTION
This script disables/deletes devices that have been inactive for a specified threshold and optionally deletes them. It generates reports and sends an email summary.

.PARAMETER SendGridApiKeyKvName
The name of the Key Vault instance containing the SendGrid API key.

.PARAMETER SendGridApiKeyKvSecretName
The name of the secret in the Key Vault containing the SendGrid API key.

.PARAMETER SendGridSenderEmailAddress
The email address of the sender.

.PARAMETER SendGridRecipientEmailAddresses
An array of recipient email addresses.

.PARAMETER SendGridApiEndpoint
The endpoint URL for the SendGrid API.

.PARAMETER DeviceDisableThreshold
Number of inactive days to determine a stale device to disable.

.PARAMETER DeviceDeleteThreshold
Number of inactive days to determine a stale device to delete.

.PARAMETER ScriptAction
The action to perform on the devices. Options are "ReportOnly", "DisableAndDelete".
#>

[CmdletBinding()]

param (
     [Parameter(Mandatory = $true)]
    [string] $SendGridApiKeyKvName,
    [Parameter(Mandatory = $true)]
    [string] $SendGridApiKeyKvSecretName,
    [Parameter(Mandatory = $true)]
    [string] $SendGridSenderEmailAddress,
    [Parameter(Mandatory = $true)]
    [string] $SendGridRecipientEmailAddresses,  # Comma-separated string
    [string] $SendGridApiEndpoint = "https://api.sendgrid.com/v3/mail/send",
    [int] $DeviceDisableThreshold = 90,
    [int] $DeviceDeleteThreshold = 180,
    [ValidateSet("ReportOnly", "DisableAndDelete")]
    [string] $ScriptAction = "ReportOnly"
)

Connect-MgGraph -Identity -NoWelcome

Disable-AzContextAutosave -Scope Process
$context = (Connect-AzAccount -Identity).context
Set-AzContext -SubscriptionName $context.Subscription -DefaultProfile $context

$ErrorActionPreference = "Stop"
$reportDir = $env:TEMP

function Send-EmailReport {
<#
.SYNOPSIS
Sends an email report using SendGrid.

.DESCRIPTION
This function sends an email report with the specified content to the given recipient email addresses using SendGrid's API.

.PARAMETER SendGridApiKey
The API key for authenticating with SendGrid.

.PARAMETER SendGridApiEndpoint
The endpoint URL for the SendGrid API.

.PARAMETER SenderEmailAddress
The email address of the sender.

.PARAMETER RecipientEmailAddresses
An array of recipient email addresses.

.PARAMETER Subject
The subject of the email.

.PARAMETER Content
The content of the email.

.PARAMETER Attachments
An array of attachments to include in the email.
#>
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

try {
    $disableDate = [datetime]::UtcNow.AddDays(-$DeviceDisableThreshold).ToString("yyyy-MM-ddTHH:mm:ssZ")
    $deleteDate = [datetime]::UtcNow.AddDays(-$DeviceDeleteThreshold).ToString("yyyy-MM-ddTHH:mm:ssZ")

    # Fetch Disabled & Deleted Devices
    $pendingDevices = Get-MgDevice -All -Filter "ApproximateLastSignInDateTime le $disableDate AND ApproximateLastSignInDateTime ge $deleteDate"
    Write-Verbose "$($pendingDevices.Count) Pending Devices to disable"

    $staleDevices = Get-MgDevice -All -Filter "ApproximateLastSignInDateTime le $deleteDate"
    Write-Verbose "$($staleDevices.Count) Stale Devices to delete"

    # Generate CSV Reports
    $pendingDevices | Export-Csv -Path "$reportDir\disabled-devices.csv" -NoTypeInformation
    $staleDevices | Export-Csv -Path "$reportDir\deleted-devices.csv" -NoTypeInformation

    # Send Email Report
    $sendGridApiKey = Get-AzKeyVaultSecret -VaultName $SendGridApiKeyKvName -Name $SendGridApiKeyKvSecretName -AsPlainText
    
    # Determine the action description based on $ScriptAction
    $actionDescription = if ($ScriptAction -eq "DisableAndDelete") {
        "The script is currently running in 'DisableAndDelete' mode and will disable and delete devices."
    } else {
        "The script is currently running in 'ReportOnly' mode and will not make any changes to the devices."
    }

    $attachments = @(
        @{
            file = "$reportDir\disabled-devices.csv"
            type = "text/csv"
        }
        @{
            file = "$reportDir\deleted-devices.csv"
            type = "text/csv"
        }
    )

    # Split comma-separated recipient addresses into an array
    $recipientArray = $SendGridRecipientEmailAddresses -split '\s*,\s*'

    Send-EmailReport -SendGridApiKey $sendGridApiKey `
                     -SendGridApiEndpoint $SendGridApiEndpoint `
                     -SenderEmailAddress $SendGridSenderEmailAddress `
                     -RecipientEmailAddresses $recipientArray `
                     -Subject "Device Cleanup Report" `
                     -Content "This is a report from MS Entra device cleanup. Pending Devices to Disable: $($pendingDevices.count), Stale Devices to Delete: $($staleDevices.count). $actionDescription" `
                     -Attachments $attachments

    # Clean Up Devices
    if ("DisableAndDelete" -eq $ScriptAction) {
        $pendingDevices | ForEach-Object {
            Write-Verbose "Disabling Device $($_.DisplayName)"
            Update-MgDevice -DeviceId $_.Id -AccountEnabled:$false
        }

        $staleDevices | ForEach-Object {
            Write-Verbose "Deleting Device $($_.DisplayName)"
            Remove-MgDevice -DeviceId $_.Id
        }
    }
}
catch {
    Write-Error $_.Exception.Message
}
