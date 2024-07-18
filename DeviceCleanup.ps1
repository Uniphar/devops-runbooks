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
    [string[]] $SendGridRecipientEmailAddresses,

    [Parameter(Mandatory = $false)]
    [string] $SendGridApiEndpoint = "https://api.sendgrid.com/v3/mail/send",

    [Parameter(Mandatory = $false)]
    [int] $DeviceDisableThreshold = 90,

    [Parameter(Mandatory = $false)]
    [int] $DeviceDeleteThreshold = 180,

    [Parameter(Mandatory = $false)]
    [ValidateSet("ReportOnly", "DisableAndDelete")]
    [string] $ScriptAction = "ReportOnly"
)

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.identity.DirectoryManagement

Disable-AzContextAutosave -Scope Process
$context = (Connect-AzAccount -Identity).context
Set-AzContext -SubscriptionName $context.Subscription -DefaultProfile $context
Connect-MgGraph -NoWelcome

$ErrorActionPreference = "Stop"
$reportDir = "c:\temp"

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

    Send-EmailReport -SendGridApiKey $sendGridApiKey `
                     -SendGridApiEndpoint $SendGridApiEndpoint `
                     -SenderEmailAddress $SendGridSenderEmailAddress `
                     -RecipientEmailAddresses $SendGridRecipientEmailAddresses `
                     -Subject "Device Cleanup Report" `
                     -Content "Pending Devices to Disable: $($pendingDevices.count), Stale Devices to Delete: $($staleDevices.count)" `
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