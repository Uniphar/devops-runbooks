<#
.SYNOPSIS
Disable user accounts on MS Entra and Active Directory based on inactivity thresholds and sends an email report.

.DESCRIPTION
The script 
 - read users from AD and MS Entra and 
 - filter users inactive in both systems
 - warn the user and manager via email in 35 days of inactivity
 - disable user in 45 days of inactivity
 - inform manager that user has been disabled
 - create and send report of disabled account

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

.PARAMETER InnactivityTime
Number of inactive days to determine a user to disable.

.PARAMETER UserWarningThreshold
Number of inactive days to determine a user to send warning.

.PARAMETER GroupId
Id of the exclusion group, members will not be evaluated incl. nested members.

.PARAMETER ADCredentials
Number of inactive days to determine a user to send warning.

.PARAMETER Testing
Boolean to indicate if the script is running in testing mode.

.PARAMETER ITSupportTeamEmailAddresses
An array of IT support team email addresses.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string] $SendGridApiKeyKvName,

    [Parameter(Mandatory = $false)]
    [string] $SendGridApiKeyKvSecretName,

    [Parameter(Mandatory = $false)]
    [string] $SendGridSenderEmailAddress,

    [Parameter(Mandatory = $false)]
    [string[]] $SendGridRecipientEmailAddresses,

    [Parameter(Mandatory = $false)]
    [string] $SendGridApiEndpoint = "https://api.sendgrid.com/v3/mail/send",

    [Parameter(Mandatory = $false)]
    [int] $InnactivityTime = 45,

    [Parameter(Mandatory = $false)]
    [int] $UserWarningThreshold = 35,

    [Parameter(Mandatory = $true)]
    [string] $groupId = "3c65e8b9-258c-4469-b0a4-18ca4c508b45",

    [Parameter(Mandatory = $false)]
    [bool] $Testing = $true,

    [Parameter(Mandatory = $true)]
    [string[]] $ITSupportTeamEmailAddresses
)

  #Get onprem AD domain admin credentials from key vault
    $domainUser = (Get-AzKeyVaultSecret -VaultName "uni-core-on-prem-kv" -Name "domain-admin-user").SecretValue #SecureString
    $domainUser = [Net.NetworkCredential]::new('', $domainUser).Password # decrypt to string
    $domainUser = -join("unipharad\", $domainUser); # add domain name to the username
    $domainPassword = (Get-AzKeyVaultSecret -VaultName "uni-core-on-prem-kv" -Name "domain-admin-pwd").SecretValue #SecureString
    $ADCredentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $domainuser,$domainPassword #combine credentials
  
    # Define $reportDir
    $reportDir = $env:TEMP


# Function to disable user in on-premises AD
function Disable-OnPremADUser {
    param (
        [string]$userPrincipalName
    )
    $user = Get-ADUser -server unidc10.uniphar.local -Credential $ADCredentials -Filter { UserPrincipalName -eq $userPrincipalName }
    if ($user) {
        Disable-ADAccount -server unidc10.uniphar.local -Credential $ADCredentials -Identity $user
        Write-Host "Disabled on-prem AD account for user: $userPrincipalName"
        return "Success"
    } else {
        Write-Host "User not found in on-prem AD: $userPrincipalName"
        return "User not found"
    }
}

# Function to disable user in Azure AD via Microsoft Graph PowerShell module
function Disable-MgUser {
    param (
        [string]$userId
    )
    try {
        Update-MgUser -UserId $userId -AccountEnabled:$false
        Write-Host "Disabled Azure AD account for user: $userId"
        return "Success"
    } catch {
        Write-Host "Failed to disable Azure AD account for user: $userId"
        return "Failed"
    }
}

function Get-GroupMembers { #recursive function to get all members of a group
    param (
        [string]$GroupId,
        [ref]$Exclusion
    )

    if (-not $GroupId) {
        Write-Error "GroupId cannot be empty."
        return
    }

    $members = Get-MgGroupMember -GroupId $GroupId -all | Select -ExpandProperty AdditionalProperties

    foreach ($member in $members) {
        if ($member.'@odata.type' -eq "#microsoft.graph.user") {
            $Exclusion.Value += $member.userPrincipalName
        }
        elseif ($member.'@odata.type' -eq "#microsoft.graph.group") {
            $group = Get-MgGroup -Filter "displayName eq '$($member.displayName)'" | Select-Object -First 1
            Get-GroupMembers -GroupId $group.Id -Exclusion $Exclusion
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
An optional array of attachments to include in the email.
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

        [Parameter(Mandatory = $false)]
        [Object[]] $Attachments
    )

    if ($Testing) {
        $date = Get-Date -Format "yyyy-MM-dd"
        $logFile = "$env:TEMP\email_log_$date.txt"
        $logContent = "To: $($RecipientEmailAddresses -join ', ')" + [Environment]::NewLine +
                      "Subject: $Subject" + [Environment]::NewLine +
                      "Content: $Content" + [Environment]::NewLine +
                      "Attachments: $($Attachments | ForEach-Object { $_.filename } -join ', ')" + [Environment]::NewLine +
                      "----------------------------------------" + [Environment]::NewLine
        Add-Content -Path $logFile -Value $logContent
    } else {
        $headers = @{
            "Authorization" = "Bearer $SendGridApiKey"
            "Content-Type"  = "application/json"
        }

        $attachments = if ($Attachments) {
            $Attachments | ForEach-Object {
                $contentCsv = Get-Content $_.file -Raw

                @{
                    content     = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($contentCsv))
                    filename    = [System.IO.Path]::GetFileName($_.file)
                    type        = $_.type
                    disposition = "attachment"
                }
            }
        } else {
            @()
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
}

# Connect to Microsoft Graph
if ($Testing) {
    Connect-MgGraph -scope User.Read.All, AuditLog.read.All, Group.Read.All -identity
} else {
    Connect-MgGraph -scope User.Read.All, AuditLog.read.All, Group.Read.All, User.ReadWrite.All -identity
}

# Get all exclusion group members' UPNs
Get-GroupMembers -GroupId $groupId -Exclusion ([ref]$exclusion)

# Gather all users in tenant
    $AllUsers = Get-MgBetaUser -Property signinactivity -all | Where-Object { $_.AccountEnabled -and $_.UserType -eq "Member" } #| Select-Object -First 500

# creating lists of active users in onprem AD - with two dates (report is $cutoffDate + disabling is $cutoffDate2)
    $activeUPNs =$null
    # Calculate the cutoff date
    $cutoffDate = (Get-Date).AddDays(-$innactivitytime)
    $cutoffDate2 = (Get-Date).AddDays(-$UserWarningThreshold)
    # Get list of all UPNs from on-prem AD that were active within the inactivity time
    $activeUsers = Get-ADUser -server unidc10.uniphar.local -Credential $ADCredentials -Filter {LastLogonDate -ge $cutoffDate} -Properties UserPrincipalName, LastLogonDate
    $activeUsers2 = Get-ADUser -server unidc10.uniphar.local -Credential $ADCredentials -Filter {LastLogonDate -ge $cutoffDate2} -Properties UserPrincipalName, LastLogonDate
    # Extract UPNs
    $activeUPNs = $activeUsers | Select-Object -ExpandProperty UserPrincipalName
    $activeUPNs2 = $activeUsers2 | Select-Object -ExpandProperty UserPrincipalName

# Create a new empty array list object for disabling
        $Report = [System.Collections.Generic.List[Object]]::new()
# Create a new empty array list object for notification
        $notification = [System.Collections.Generic.List[Object]]::new()
   
Foreach ($user in $AllUsers) {
    # Null variables
        $SignInActivity = $null
        $Licenses = $null
        $Manager = $null
        $ManagerEmail = $null
        $maxdate = $null
        $CreatedDateTime = $null
       
    
    # Display progress output 
       Write-host "Gathering sign-in information for $($user.DisplayName)" -ForegroundColor Cyan


    # Count the last signing date from all posible variants
           # Retrieve the date values
            $LastInteractiveSignIn = $user.SignInActivity.LastSignInDateTime
            $LastNonInteractiveSignin = $user.SignInActivity.LastNonInteractiveSignInDateTime
            $LastSuccessfullSignInDate = $user.SignInActivity.LastSuccessfulSignInDateTime

            $maxdate = $null

            if ($LastInteractiveSignIn -ne $null) {
                $maxdate = $LastInteractiveSignIn
            }

            if ($LastNonInteractiveSignin -ne $null -and ($maxdate -eq $null -or $LastNonInteractiveSignin -gt $maxdate)) {
                $maxdate = $LastNonInteractiveSignin
            }

            if ($LastSuccessfullSignInDate -ne $null -and ($maxdate -eq $null -or $LastSuccessfullSignInDate -gt $maxdate)) {
                $maxdate = $LastSuccessfullSignInDate
            }
    #reporting on the screen and get $daysInactive for user
    if ($maxdate) {
        $daysInactive = [math]::Round(((Get-Date) - $maxdate).TotalDays)
        Write-Host "Last sign in date is $($maxdate), Days of inactivity: $daysInactive" -ForegroundColor Blue
    } else {
        Write-Host "Last sign in date is not available" -ForegroundColor Cyan
    }
     # Retrieve account creation date
    $accountCreationDate = $user.CreatedDateTime
        if (-not $accountCreationDate) {
            $accountCreationDate = [datetime]"1/1/2000" #put some old date if it is empty
        }
    $daysSinceCreation = [math]::Round(((Get-Date) - $accountCreationDate).TotalDays)

if ($daysSinceCreation -gt 21 -and $maxDate -lt (Get-Date).AddDays(-$InnactivityTime)) { # if user inactive and the account is not new, then 
        # Get current user license information
        $licenses = (Get-MgBetaUserLicenseDetail -UserId $user.id).SkuPartNumber -join ", "
    
        # Proceed silently to get manager information
            try {
                $managerid = Get-MgUserManager -UserId $user.id -ErrorAction SilentlyContinue

                if ($managerid) {
                $manager = get-mguser -UserId $managerid.Id -ErrorAction SilentlyContinue
                $managerEmail = $manager.Mail 
                }
            } catch {
                $manager = $null
                $managerEmail = $null
            }
    
      # Verify if the user is in the $activeUPNs list, if not, continue
      if ($user.UserPrincipalName -notin $activeUPNs) {
    
      # Verify if the user is in the $exclusion list, if not, continue
      if ($user.UserPrincipalName -notin $exclusion) {
    
    
        # Create informational object to add to report
        $obj1 = [pscustomobject][ordered]@{
            DisplayName                = $user.DisplayName
            UserPrincipalName          = $user.UserPrincipalName
            Email                      = $user.Mail
            Manager                    = $Manager.DisplayName
            ManagerEmail               = $manager.Mail
            Licenses                   = $licenses
            Company                    = $user.CompanyName
            CreatedDateTime            = $User.CreatedDateTime
            LastActivityDate           = $maxDate
          }
    Write-host "Adding user to disabling list $($user.DisplayName)" -ForegroundColor Red
    # Add current user info to report
    $report.Add($obj1)
    }}
}



    #create a report for users to send warning 
if ($daysInactive -eq $UserWarningThreshold) { # if user inactive then 
            
    # Proceed silently to get manager information
            try {
                $managerid = Get-MgUserManager -UserId $user.id -ErrorAction SilentlyContinue

                if ($managerid) {
                $manager = get-mguser -UserId $managerid.Id -ErrorAction SilentlyContinue
                $managerEmail = $manager.Mail 
                }
            } catch {
                $manager = $null
                $managerEmail = $null
            }
    
      # Verify if the user is in the $activeUPNs list, if not, continue
      if ($user.UserPrincipalName -notin $activeUPNs2) {
    
      # Verify if the user is in the $exclusion list, if not, continue
if ($user.UserPrincipalName -notin $exclusion) {
    
        # Create informational object to add to report
        $obj2 = [pscustomobject][ordered]@{
            DisplayName                = $user.DisplayName
            UserPrincipalName          = $user.UserPrincipalName
            Email                      = $user.Mail
            Manager                    = $Manager.DisplayName
            ManagerEmail               = $manager.Mail
            Licenses                   = $licenses
            Company                    = $user.CompanyName
            CreatedDateTime            = $User.CreatedDateTime
            LastActivityDate           = $maxDate
          }
    Write-host "Adding user to notification list $($user.DisplayName)" -ForegroundColor Yellow
    # Add current user info to report
    $notification.Add($obj2)
    }}
}

}

$report | Export-CSV -path "$reportDir\disabled_users.csv" -NoTypeInformation
$notification | Export-CSV -path "$reportDir\notification_list.csv" -NoTypeInformation

# Initialize report array
$disableReport = @()

# Iterate over the report and disable users
foreach ($user in $report) {
    if (-not $Testing) {
        $onPremResult = Disable-OnPremADUser -userPrincipalName $user.UserPrincipalName
        $azureResult = Disable-MgUser -userId $user.UserPrincipalName
    } else {
        $onPremResult = "Testing mode - no action taken"
        $azureResult = "Testing mode - no action taken"
    }

    # Add result to report
    $disableReport += [PSCustomObject]@{
        UserPrincipalName = $user.UserPrincipalName
        OnPremResult     = $onPremResult
        AzureResult      = $azureResult
    }
}

# Export the report to CSV
$disableReport | Export-Csv -Path "$reportDir\DisableReport.csv" -NoTypeInformation

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
)

Send-Email -SendGridApiKey $SendGridApiKey `
           -SendGridApiEndpoint $SendGridApiEndpoint `
           -SenderEmailAddress $SendGridSenderEmailAddress `
           -RecipientEmailAddresses $ITSupportTeamEmailAddresses `
           -Subject "Inactive users report" `
           -Content "Users disabled: $($report.count), Users notified that account will be disabled: $($notification.count)" `
           -Attachments $attachments

# Iterate over the report and send notification to user and manager 10 days before disabling
foreach ($user in $notification) {
    $SendGridRecipientEmailAddresses = @($user.Email, $user.ManagerEmail)
    Send-Email -SendGridApiKey $SendGridApiKey `
               -SendGridApiEndpoint $SendGridApiEndpoint `
               -SenderEmailAddress $SendGridSenderEmailAddress `
               -RecipientEmailAddresses $SendGridRecipientEmailAddresses `
               -Subject "User $($user.DisplayName) will be disabled" `
               -Content "User $($user.DisplayName) will be disabled in 10 days because of inactivity. User must login to his account to ensure the account stays enabled. In case of any issues please create a support request. This is an automated email, please do not answer."
}
