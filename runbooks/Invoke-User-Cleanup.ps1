<#
.SYNOPSIS
Disable member accounts on MS Entra based on inactivity thresholds and sends an email report.

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
    [string] $groupId = "3c65e8b9-258c-4469-b0a4-18ca4c508b45"

   )

  #Get onprem AD domain admin credentials from key vault
    $domainUser = (Get-AzKeyVaultSecret -VaultName "uni-core-on-prem-kv" -Name "domain-admin-user").SecretValue #SecureString
    $domainUser = [Net.NetworkCredential]::new('', $domainUser).Password # decrypt to string
    $domainuser = -join("unipharad\", $domainuser); # add domain name to the username
    $domainPassword = (Get-AzKeyVaultSecret -VaultName "uni-core-on-prem-kv" -Name "domain-admin-pwd").SecretValue #SecureString
    $ADCredentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $domainuser,$domainPassword #combine credentials
  
# Function to disable user in on-premises AD
function Disable-OnPremADUser {
    param (
        [string]$userPrincipalName
    )
    $user = Get-ADUser -Credential $ADCredentials -Filter { UserPrincipalName -eq $userPrincipalName }
    if ($user) {
        Disable-ADAccount -Credential $ADCredentials -Identity $user
        Write-Host "Disabled on-prem AD account for user: $userPrincipalName"
        return "Success"
    } else {
        Write-Host "User not found in on-prem AD: $userPrincipalName"
        return "User not found"
    }
}

# Function to disable user in Azure AD via MS Graph PowerShell module
function Disable-AzureADUser {
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


# Function to get all members of a exclusion group, including nested groups
function Get-GroupMembers {
    param (
        [string]$GroupId
    )

    # Get direct members of the group
    $members = Get-AzureADGroupMember  -ObjectId $GroupId -all $true

    foreach ($member in $members) {
        if ($member.ObjectType -eq "User") {
            # Output the UPN of the user
            $member.UserPrincipalName
        } elseif ($member.ObjectType -eq "Group") {
            # Recursively get members of the nested group
            Get-GroupMembers -GroupId $member.ObjectId
        }
    }
}

#function to sends an email using SendGrid.
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

# Function to send an email using SendGrid without attachments.
function Send-EmailNotification {
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
        [string] $Subject,

        [Parameter(Mandatory = $true)]
        [string] $Content
    )

    # Construct the email payload
    $emailPayload = @{
        personalizations = @(
            @{
                to = $RecipientEmailAddresses | ForEach-Object { @{ email = $_ } }
                subject = $Subject
            }
        )
        from = @{
            email = $SenderEmailAddress
        }
        content = @(
            @{
                type = "text/plain"
                value = $Content
            }
        )
    }

    # Convert the payload to JSON
    $jsonPayload = $emailPayload | ConvertTo-Json -Depth 10

    # Send the email using SendGrid API
    $response = Invoke-RestMethod -Method Post -Uri $SendGridApiEndpoint -Headers @{
        "Authorization" = "Bearer $SendGridApiKey"
        "Content-Type"  = "application/json"
    } -Body $jsonPayload

    return $response
}

# Connect to Azure AD
Connect-AzureAD -identity

# Get all members' UPNs
$exclusion = Get-GroupMembers -GroupId $groupId | ForEach-Object { $_ }

# Connect to Microsoft Graph
Connect-MgGraph -scope User.Read.All, AuditLog.read.All, Group.Read.All -identity
#scope for disabling - needs to be switched for disabling !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
#Connect-MgGraph -scope User.Read.All, AuditLog.read.All, Group.Read.All, User.ReadWrite.All -identity

# Gather all users in tenant
    $AllUsers = Get-MgBetaUser -Property signinactivity -all | Where-Object { $_.AccountEnabled -and $_.UserType -eq "Member" } #| Select-Object -First 500

# creating list of active users in onprem AD - for two options
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


    # Count the last signing date from all posiible variants
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

if ($daysSinceCreation -gt 21 -and $maxDate -lt (Get-Date).AddDays(-$innactivitytime)) { # if user inactive and the account is not new, then 
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




$report | Export-CSV -path C:\temp\disabled_users.csv -NoTypeInformation
$notification | Export-CSV -path C:\temp\notification_list.csv -NoTypeInformation


# Initialize report array
$disableReport = @()

# Iterate over the report and disable users
foreach ($user in $report) {
    $onPremResult = Disable-OnPremADUser -userPrincipalName $user.UserPrincipalName
    $azureResult = Disable-AzureADUser -userId $user.Id

    # Add result to report
    $disableReport += [PSCustomObject]@{
        UserPrincipalName = $user.UserPrincipalName
        OnPremResult     = $onPremResult
        AzureResult      = $azureResult
    }
}

# Export the report to CSV
$disableReport | Export-Csv -Path "C:\Temp\DisableReport.csv" -NoTypeInformation


#Send report after disabling to IT support teams

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

    Send-EmailReport -SendGridApiKey $sendGridApiKey `
                     -SendGridApiEndpoint $SendGridApiEndpoint `
                     -SenderEmailAddress $SendGridSenderEmailAddress `
                     -RecipientEmailAddresses $SendGridRecipientEmailAddresses `
                     -Subject "Inactive users report" `
                     -Content "Users disabled: $($disabled_users.count), Users notified that account will be disabled: $($notification.count)" `
                     -Attachments $attachments

# Iterate over the report and sendNotification to user and manager 10 days before disabling
foreach ($user in $report) {
$SendGridRecipientEmailAddresses = "$userEmail,$userManagerEmail"
    Send-Emailnotification -SendGridApiKey $sendGridApiKey `
                             -SendGridApiEndpoint $SendGridApiEndpoint `
                             -SenderEmailAddress $SendGridSenderEmailAddress `
                             -RecipientEmailAddresses $SendGridRecipientEmailAddresses `
                             -Subject "User $user.DisplayName will be disabled" `
                             -Content "User $user.DisplayName will be disabled in 10 days because of inactivity. User must login to his account to ensure the account stays enabled. In case on any issues please create a support request."

}
