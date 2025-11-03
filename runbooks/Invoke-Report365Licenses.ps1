# RUNBOOK (Azure Automation compatible)
# This runbook gathers Microsoft 365 license (SKU) information and produces a CSV and HTML report.
# Designed to run in Azure Automation using the Automation Account's managed identity.
# Behavior summary:
# - Authenticates to Microsoft Graph using the Automation Account managed identity (Connect-MgGraph -Identity).
# - Retrieves subscribed SKUs (Get-MgSubscribedSku) and generates a CSV + HTML report saved to the runbook temp folder.
# - Sends immediate alert emails via SendGrid when configured monitored SKUs fall below a configurable threshold.
# - Sends the full CSV+HTML report as an email attachment automatically on the 1st day of each month.
# - Uses Azure Key Vault to retrieve the SendGrid API key (the managed identity must have access to the Key Vault secret).
# Notes for deployment:
# - Ensure the Microsoft.Graph.* modules are installed in the Automation Account modules gallery.
# - Grant the Automation Account managed identity the following Microsoft Graph Application permissions and grant admin consent:
#     Directory.Read.All, Organization.Read.All
# - Grant the managed identity access to the Key Vault secret containing the SendGrid API key (Get/List secret permissions or appropriate IAM).
# - Configure the runbook schedule (daily) in Azure Automation; the runbook will automatically send monthly full reports on day 1.



param (
    [Parameter(Mandatory = $true)] [string]$KeyVaultName,
    [Parameter(Mandatory = $true)] [string]$SendGridSecretName,
    [Parameter(Mandatory = $true)] [string]$FromEmail,
    [Parameter(Mandatory = $true)] [string]$ToEmail,
    [Parameter(Mandatory = $false)] [string]$MonitorSkuId1 = "",
    [Parameter(Mandatory = $false)] [string]$MonitorSkuId2 = "",
    [Parameter(Mandatory = $false)] [string]$MonitorSkuId3 = "",
    [Parameter(Mandatory = $false)] [string]$MonitorSkuId4 = "",
    [Parameter(Mandatory = $false)] [string]$MonitorSkuId5 = "",
    [Parameter(Mandatory = $false)] [int]$MinimumLicenseThreshold = 5
)

# Use Azure Automation temp path ($env:TEMP) when available; fallback to C:\temp for local testing.
if ($env:AZUREPS_HOST_ENVIRONMENT) {
    # In Azure Automation, $env:TEMP is the correct sandbox temporary path.
    $TempPath = $env:TEMP
}
else {
    # Local testing fallback
    $TempPath = "C:\\temp"
}
if (!(Test-Path -Path $TempPath)) {
    Write-Host "Creating temporary directory: $TempPath"
    New-Item -Path $TempPath -ItemType Directory | Out-Null
}

[string]$RunDate = Get-Date -format "dd-MMM-yyyy HH:mm:ss"
$CSVOutputFile = Join-Path -Path $TempPath -ChildPath "Microsoft365LicenseServicePlans.csv"
$ReportFile = Join-Path -Path $TempPath -ChildPath "Microsoft365LicenseServicePlans.html"

# --- SendGrid and Key Vault integration ---
# Retrieve the SendGrid API key from Azure Key Vault using the Automation Account managed identity.
# The managed identity requires Key Vault access (Get secrets) or an appropriate role assignment.
Write-Output "========================================"
Write-Output "KEY VAULT ACCESS DEBUG INFORMATION:"
Write-Output "========================================"
Write-Output "Attempting to retrieve SendGrid secret from Key Vault (managed identity will be used):"
Write-Output "Key Vault Name: $KeyVaultName"
Write-Output "Secret Name: $SendGridSecretName"

Connect-AzAccount -Identity -ErrorAction Stop
Write-Output "✓ Connected to Azure (Az) via managed identity"

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
}
catch {
    Write-Output "✗ Failed to retrieve SendGrid API Key from Key Vault '$KeyVaultName'"
    Write-Output "Error Type: $($_.Exception.GetType().FullName)"
    Write-Output "Error Message: $($_.Exception.Message)"
    $SendGridApiKey = ""
}
Write-Output "========================================"

# Connect to Microsoft Graph using the Automation Account managed identity.
Write-Host "Connecting to Microsoft Graph using Automation Managed Identity..."

## Required application permissions (grant these to the Automation Account service principal and consent):
## - Directory.Read.All (Application): read subscribed SKUs and directory information
## - Organization.Read.All (Application): read organization details used by Get-MgOrganization
# Note: when using -Identity (managed identity) do NOT pass -Scopes; passing -Scopes with -Identity causes a ParameterSet ambiguity error.

Write-Output "DEBUG: About to call Connect-MgGraph -Identity"
try {
    Connect-MgGraph -Identity -NoWelcome -ErrorAction Stop
    Write-Output "✓ Connected to Microsoft Graph via Managed Identity"
}
catch {
    Write-Output "✗ FAILED to connect to Microsoft Graph"
    Write-Output "Error Type: $($_.Exception.GetType().FullName)"
    Write-Output "Error Message: $($_.Exception.Message)"
    Write-Output "Stack Trace: $($_.ScriptStackTrace)"
    throw
}

# Verify connection and permissions
Write-Output "========================================"
Write-Output "MICROSOFT GRAPH CONNECTION DEBUG:"
Write-Output "========================================"
try {
    $mgContext = Get-MgContext
    Write-Output "✓ Graph Context Retrieved:"
    Write-Output "  Account: $($mgContext.Account)"
    Write-Output "  AppName: $($mgContext.AppName)"
    Write-Output "  TenantId: $($mgContext.TenantId)"
    Write-Output "  Scopes: $($mgContext.Scopes -join ', ')"
    Write-Output "  AuthType: $($mgContext.AuthType)"
    Write-Output "  ContextScope: $($mgContext.ContextScope)"
    
    # Check if we have app-level permissions (not delegated)
    if ($mgContext.AuthType -eq 'AppOnly') {
        Write-Output "  ✓ Using App-Only (Application) permissions (suitable for runbook)"
    }
    elseif ($mgContext.AuthType -eq 'Delegated') {
        Write-Output "  ⚠ Using Delegated permissions (may cause access issues in runbook)"
    }
} catch {
    Write-Output "✗ Could not get Graph context: $($_.Exception.Message)"
    Write-Output "Stack Trace: $($_.ScriptStackTrace)"
}
Write-Output "========================================"

# Built-in friendly names for common licenses (fallback if Microsoft data unavailable). These are used
# only when the Microsoft product names CSV cannot be downloaded or parsed.
$BuiltInSkuNames = @{
    # Microsoft 365 Core Licenses
    '05e9a617-0261-4cee-bb44-138d3ef5d965' = 'Microsoft 365 E3'
    '06ebc4ee-1bb5-47dd-8120-11324bc54e06' = 'Microsoft 365 E5'
    '66b55226-6b4f-492c-910c-a3b7a3c9d993' = 'Microsoft 365 F3'
    '50f60901-3181-4b75-8a2c-4c8e4c1d5a72' = 'Microsoft 365 F1'
    '639dec6b-bb19-468b-871c-c5c441c4b0cb' = 'Microsoft 365 Copilot'
    
    # Power BI
    'a403ebcc-fae0-4ca2-8c8c-7a907fd6c235' = 'Power BI Pro'
    'f30db892-07e9-47e9-837c-80727f46fd3d' = 'Power BI Premium Per User'
    '7b26f5ab-a763-4c00-a1ac-f6c4b5506945' = 'Power BI Premium P1 Add-On'
    '8ecbd3c1-b108-437c-a859-e3c125e3f83f' = 'Power BI Premium EM2 Add-On'
    'c1d032e0-5619-4761-9b5c-75b6831e1711' = 'Power BI Premium Per User'
    
    # Power Platform
    'b30411f5-fea1-4a59-9ad9-3db7c7ead579' = 'Power Apps Per User'
    '4a51bf65-409c-4a91-b845-1121b571cc9d' = 'Power Automate Per User'
    '87bbbc60-4754-4998-8c88-227dca264858' = 'Power Apps Individual User'
    'dcb1a3ae-b33f-4487-846a-a640262fadf4' = 'Power Apps Viral'
    'bf666882-9c9b-4b2e-aa2f-4789b0a52ba2' = 'Power Apps Per App'
    'b4d7b828-e8dc-4518-91f9-e123ae48440d' = 'Power Apps Per App'
    '5b631642-bd26-49fe-bd20-1daaa972ef80' = 'Power Apps for Developer'
    'b3a42176-0a8c-4c3f-ba4e-f2b37fe5be6b' = 'Power Automate Business Process'
    '253ce8d3-6122-4240-8b04-f434a8fa831f' = 'Power Automate Per Process'
    'eda1941c-3c4f-4995-b5eb-e85a42175ab9' = 'Power Automate Attended RPA'
    '3f9f06f5-3c31-472c-985f-62d9c10ec167' = 'Power Pages vTrial for Makers'
    'debc9e58-f2d7-412c-a0b6-575608564228' = 'Power Pages Authenticated Users T1 (100/site/month)'
    '74a8790a-f2e7-41c4-98bf-dc25bf252af0' = 'Power Pages Anonymous Users T1 (500/site/month)'
    
    # Security & Compliance
    '8c4ce438-32a7-4ac5-91a6-e22ae08d9c8b' = 'Rights Management Adhoc'
    '093e8d14-a334-43d9-93e3-30589a8b47d0' = 'Rights Management Basic'
    '3dd6cf57-d688-4eed-ba52-9e40b5468c3e' = 'Threat Intelligence'
    '84a661c4-e949-4bd2-a560-ed7766fcaf2b' = 'Microsoft Entra ID P2'
    '4ef96642-f096-40de-a3e9-d83fb2f90211' = 'Microsoft Defender for Office 365 (Plan 1)'
    'f9602137-2203-447b-9fff-41b36e08ce5d' = 'Microsoft Entra Suite'
    '555af716-7534-4f72-a79c-d5a421dd3c5c' = 'Microsoft Entra Private Access Premium'
    '26124093-3d78-432b-b5dc-48bf992543d5' = 'Microsoft 365 E5 Security'
    
    # Microsoft Stream & Storage
    '1f2f344a-700d-42c9-9427-5cea1d5d7ba6' = 'Microsoft Stream'
    '6470687e-a428-4b7a-bef2-8a291ad947c9' = 'Windows Store for Business'
    '99049c9c-6011-4908-bf17-15f496e6519d' = 'SharePoint Storage'
    '1fc08a02-8b3d-43b9-831e-f76859e04e1a' = 'SharePoint Standard'
    
    # Telephony & Communication
    '440eaaa8-b3e0-484b-a8be-62870b9ba70a' = 'Phone System - Virtual User'
    'e43b5b99-8dfb-405f-9987-dc307f34bcbd' = 'Microsoft 365 Phone System'
    '0dab259f-bf13-4952-b7f8-7db8f131b28d' = 'Domestic Calling Plan'
    'd3b4fe1f-9992-4930-8acb-ca6ec609365e' = 'Domestic and International Calling Plan'
    '47794cd0-f0e5-45c5-9033-2eb6b5fc84e0' = 'Communications Credits'
    '1c27243e-fb4d-42b1-ae8c-fe25c9616588' = 'Microsoft Teams Audio Conferencing (Select Dial Out)'
    
    # Microsoft Teams
    '36a0f3b3-adb5-49ea-bf66-762134cf063a' = 'Microsoft Teams Premium'
    '4cde982a-ede4-4409-9ae6-b003453c8ea6' = 'Microsoft Teams Rooms Pro'
    
    # Dynamics 365 - Customer Engagement
    '749742bf-0d37-4158-a120-33567104deeb' = 'Dynamics 365 Customer Service Enterprise'
    '1e1a282c-9c54-43a2-9310-98ef728faace' = 'Dynamics 365 Sales Enterprise'
    '2edaa1dc-966d-4475-93d6-8ee8dfd96877' = 'Dynamics 365 Sales Premium'
    '3489e6e2-6bfe-409a-b6f4-1ee106f2db5b' = 'Dynamics 365 Sales Insights'
    '1e615a51-59db-4807-9957-aa83c3657351' = 'Dynamics 365 Customer Service Enterprise vTrial'
    '6ec92958-3cc1-49db-95bd-bc6b3798df71' = 'Dynamics 365 Sales Premium vTrial'
    'c7d15985-e746-4f01-b113-20b575898250' = 'Dynamics 365 Field Service'
    'eb18b715-ea9d-4290-9994-2ebf4b5042d2' = 'Dynamics 365 Customer Service Enterprise Attach'
    'ff22b8d4-5073-4b24-ba45-84ad5d9b6642' = 'Dynamics 365 Customer Insights Attach'
    '84a1cdd0-11b5-4b66-8f6b-b27d385ce1bc' = 'Dynamics 365 Customer Service Messaging'
    '977464c4-bfaf-4b67-b761-a9bb735a2196' = 'Dynamics 365 Customer Service Auto Routing Add-On'
    '606b54a9-78d8-4298-ad8b-df6ef4481c80' = 'Dynamics 365 Chatbots vTrial'
    
    # Dynamics 365 - Finance & Operations
    '55c9eb4e-c746-45b4-b255-9ab6b19d5c62' = 'Dynamics 365 Finance'
    'fe896f91-8c82-4007-94ef-b46e19e50bfa' = 'Dynamics 365 Finance Premium'
    'd721f2e4-099b-4105-b40e-872e46cad402' = 'Dynamics 365 Finance Attach'
    'aebb5c59-2f81-492b-8b0b-fb8788f2be9c' = 'Dynamics 365 Operations Order Lines'
    'fcecd1f9-a91e-488d-a918-a96cdb6ce2b0' = 'Dynamics 365 for Operations Trial'
    'b75074f1-4c54-41bf-970f-c9ac871567f5' = 'Dynamics 365 Operations Activity'
    '090b4a96-8114-4c95-9c91-60e81ef53302' = 'Dynamics 365 Supply Chain Management Attach'
    '3bbd44ed-8a70-4c07-9088-6232ddbd5ddd' = 'Dynamics 365 for Operations Devices'
    'e485d696-4c87-4aac-bf4a-91b2fb6f0fa7' = 'Dynamics 365 Operations Sandbox Tier 2'
    '28cea2ad-3802-496a-a44e-517b37baee5b' = 'Dynamics 365 Operations Enterprise Storage'
    'ac4f5985-fa17-431f-882d-607254fe82fd' = 'Dynamics 365 Operations Enterprise Storage File'
    '673afb9d-d85b-40c2-914e-7bf46cd5cd75' = 'Dynamics 365 Asset Management'
    'c595cac9-ee7e-429b-9f8b-6cf377d2e7f5' = 'Dynamics 365 Globalization E-Invoicing'
    
    # Dynamics 365 - Human Resources
    '941a27e3-820c-47a0-9d4b-5c28088939c8' = 'Dynamics 365 Human Resources'
    'acf4f594-ff13-4fde-ba8f-2d7d72a7aafa' = 'Dynamics 365 Human Resources Self Service'
    'dda540bf-0ad9-4f1c-afbf-ee1495c40127' = 'Dynamics 365 Human Resources Sandbox'
    '3a256e9a-15b6-4092-b0dc-82993f4debc6' = 'Dynamics 365 for HCM Trial'
    
    # Dynamics 365 - Business Central
    '2880026b-2b0c-4251-8656-5d41ff11e3aa' = 'Dynamics 365 Business Central Essentials'
    '9a1e33ed-9697-43f3-b84c-1b0959dbb1d4' = 'Dynamics 365 Business Central for Accountants'
    '57740eb8-785c-411d-8446-19d4cd1909c0' = 'Dynamics 365 Business Central Device'
    '2e3c4023-80f6-4711-aa5d-29e0ecb46835' = 'Dynamics 365 Business Central Team Member'
    '1d506c23-1702-46f1-b940-160c55f98d05' = 'Dynamics 365 Business Central Essentials Attach'
    
    # Dynamics 365 - Other
    '98619618-9dc8-48c6-8f0c-741890ba5f93' = 'Dynamics 365 Project Operations'
    '7ac9fe77-66b7-4e5e-9e46-10eed1cff547' = 'Dynamics 365 Team Members'
    'e77d538c-4ebd-4118-8d7c-f021110424bf' = 'Microsoft Cloud for Healthcare vTrial'
    
    # Other Microsoft Services
    'c5928f49-12ba-48f7-ada3-0d743a3601d5' = 'Visio Plan 2'
    '53818b1b-4a27-454b-8896-0dba576410e6' = 'Project Plan 3'
    '2b317a4a-77a6-4188-9437-b68a77b4e2c6' = 'Intune Device'
    '19ec0d23-8335-4cbd-94ac-6050e30712fa' = 'Exchange Online Plan 2'
    'a2367322-2be4-443f-837c-06798507b89d' = 'Remote Help Add-On'
    '726a0894-2c77-4d65-99da-9775ef05aad1' = 'Microsoft Business Center'
    
    # Dataverse & CDS
    'e612d426-6bc3-4181-9658-91aa906b0ac0' = 'Dataverse Database Capacity'
    'd2dea78b-507c-4e56-b400-39447f4738f8' = 'Dataverse AI Capacity'
}

# Get the basic information about tenant subscriptions
Write-Output "========================================"
Write-Output "RETRIEVING SUBSCRIBED SKUs:"
Write-Output "========================================"
Write-Output "DEBUG: Calling Get-MgSubscribedSku (requires Directory.Read.All application permission)"
try {
    [array]$Skus = Get-MgSubscribedSku -ErrorAction Stop
    Write-Output "✓ Successfully retrieved $($Skus.Count) SKUs"
} catch {
    Write-Output "✗ FAILED to retrieve SKUs - This is likely a permissions issue"
    Write-Output "Error Type: $($_.Exception.GetType().FullName)"
    Write-Output "Error Message: $($_.Exception.Message)"
    Write-Output "Stack Trace: $($_.ScriptStackTrace)"
    Write-Output ""
    Write-Output "TROUBLESHOOTING:"
    Write-Output "1. Ensure the Automation Account managed identity has 'Directory.Read.All' application permission"
    Write-Output "2. Admin consent must be granted for the permission (not just granted via IAM)"
    Write-Output "3. The permission must be at APPLICATION LEVEL, not delegated"
    Write-Output "4. Grant permission via: Entra ID > App registrations > (your app) > API permissions"
    Write-Output "5. Verify: Entra ID > Enterprise applications > (your app) > Permissions > Admin consent status = 'Granted'"
    Write-Output ""
    throw
}
Write-Output "========================================"



# It's used to resolve SKU and service plan code names to human-friendly values
Write-Output "========================================"
Write-Output "PRODUCT NAMES DOWNLOAD:"
Write-Output "========================================"
Write-Output "Attempting to download latest product names from Microsoft (with short timeout)..."
$ProductDataAvailable = $false
$ProductNamesFile = Join-Path -Path $TempPath -ChildPath "ProductNames.csv"

# Try to download the latest product names CSV from Microsoft (optional, non-blocking)
Try {
    # Microsoft's official Product names and service plan identifiers for licensing
    # This URL is updated regularly by Microsoft
    $DirectCsvUrl = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
    
    Write-Output "DEBUG: Starting direct download from: $DirectCsvUrl"
    Write-Output "DEBUG: Current time: $(Get-Date -Format 'o')"
    Write-Output "DEBUG: Timeout set to 10 seconds"
    
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Output "DEBUG: Invoking Invoke-WebRequest..."
    $ProductInfoRequest = Invoke-WebRequest -Uri $DirectCsvUrl -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
    $sw.Stop()
    Write-Output "DEBUG: Download completed in $($sw.ElapsedMilliseconds)ms"
    
    Write-Output "DEBUG: Response received, size: $($ProductInfoRequest.Content.Length) bytes"
    Write-Output "DEBUG: Parsing CSV content using streaming parser..."
    $sw.Restart()
    
    # Use streaming parser instead of ConvertFrom-Csv for large files (avoids memory issues in sandbox)
    $csvContent = $ProductInfoRequest.Content
    $lines = $csvContent -split "`n" | Where-Object { $_.Trim().Length -gt 0 }
    
    # Find the header line (skip empty lines and find first line with reasonable column count)
    $headerLineIndex = 0
    $headerLine = $null
    $headers = @()
    for ($h = 0; $h -lt [Math]::Min(10, $lines.Count); $h++) {
        $testHeaders = $lines[$h] -split ',' | ForEach-Object { $_.Trim('"').Trim() }
        if ($testHeaders.Count -ge 3) {  # Real header should have at least 3 columns
            $headerLine = $lines[$h]
            $headers = $testHeaders
            $headerLineIndex = $h
            break
        }
    }
    
    if ($headers.Count -lt 3) {
        throw "Could not find valid CSV header (need at least 3 columns)"
    }
    
    Write-Output "DEBUG: Found $($lines.Count) lines, $(($headers).Count) columns"
    Write-Output "DEBUG: First header: $($headers[0])"
    
    $ProductData = @()
    for ($i = $headerLineIndex + 1; $i -lt $lines.Count; $i++) {
        if (($i % 500) -eq 0) {
            Write-Output "DEBUG: Parsed $i rows so far..."
        }
        $line = $lines[$i]
        if ($line.Trim().Length -eq 0) { continue }
        
        # Simple CSV parser - split on comma, handling quoted fields
        $values = @()
        $current = ""
        $inQuotes = $false
        for ($j = 0; $j -lt $line.Length; $j++) {
            $char = $line[$j]
            if ($char -eq '"') {
                $inQuotes = -not $inQuotes
            }
            elseif ($char -eq ',' -and -not $inQuotes) {
                $values += $current.Trim('"').Trim()
                $current = ""
            }
            else {
                $current += $char
            }
        }
        $values += $current.Trim('"').Trim()
        
        # Create object from headers and values
        $obj = [PSCustomObject]@{}
        for ($k = 0; $k -lt $headers.Count; $k++) {
            $obj | Add-Member -NotePropertyName $headers[$k] -NotePropertyValue $values[$k]
        }
        $ProductData += $obj
    }
    
    $sw.Stop()
    Write-Output "DEBUG: CSV parsed in $($sw.ElapsedMilliseconds)ms, row count: $($ProductData.Count)"
    
    Write-Output "DEBUG: Exporting to cache file: $ProductNamesFile"
    $sw.Restart()
    $ProductData | Export-Csv -Path $ProductNamesFile -NoTypeInformation -Encoding UTF8
    $sw.Stop()
    Write-Output "DEBUG: Export completed in $($sw.ElapsedMilliseconds)ms"
    
    Write-Output "✓ Product names downloaded and cached successfully to: $ProductNamesFile"
    $ProductDataAvailable = $true
    
}
Catch {
    Write-Output "✗ Could not download from direct URL"
    Write-Output "Error Type: $($_.Exception.GetType().FullName)"
    Write-Output "Error Message: $($_.Exception.Message)"
    Write-Output "Error Details: $($_ | Out-String)"
    
    # Fallback: Try to find the CSV link from the documentation page
    Try {
        Write-Output "DEBUG: Attempting fallback method - fetching documentation page"
        Write-Output "DEBUG: Current time: $(Get-Date -Format 'o')"
        $LicensingPageUrl = "https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference"
        
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        Write-Output "DEBUG: Invoking Invoke-WebRequest for documentation page..."
        $LicensingPageRequest = Invoke-WebRequest -Uri $LicensingPageUrl -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
        $sw.Stop()
        Write-Output "DEBUG: Documentation page downloaded in $($sw.ElapsedMilliseconds)ms"
        
        Write-Output "DEBUG: Searching for CSV link in page..."
        $DownloadLink = ($LicensingPageRequest.Links | Where-Object { $_.href -like '*.csv' }).href
        Write-Output "DEBUG: Found $(@($DownloadLink).Count) CSV links"
        
        If ($DownloadLink) {
            # Make sure the link is absolute
            If ($DownloadLink -notlike "http*") {
                $DownloadLink = "https://learn.microsoft.com$DownloadLink"
            }
            
            Write-Output "DEBUG: CSV link found: $DownloadLink"
            Write-Output "DEBUG: Downloading CSV from link..."
            $sw.Restart()
            $ProductInfoRequest = Invoke-WebRequest -Uri $DownloadLink -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
            $sw.Stop()
            Write-Output "DEBUG: Downloaded in $($sw.ElapsedMilliseconds)ms"
            
            Write-Output "DEBUG: Parsing CSV using streaming parser..."
            $sw.Restart()
            
            # Use streaming parser instead of ConvertFrom-Csv for large files
            $csvContent = $ProductInfoRequest.Content
            $lines = $csvContent -split "`n" | Where-Object { $_.Trim().Length -gt 0 }
            
            # Find the header line (skip empty lines and find first line with reasonable column count)
            $headerLineIndex = 0
            $headerLine = $null
            $headers = @()
            for ($h = 0; $h -lt [Math]::Min(10, $lines.Count); $h++) {
                $testHeaders = $lines[$h] -split ',' | ForEach-Object { $_.Trim('"').Trim() }
                if ($testHeaders.Count -ge 3) {  # Real header should have at least 3 columns
                    $headerLine = $lines[$h]
                    $headers = $testHeaders
                    $headerLineIndex = $h
                    break
                }
            }
            
            if ($headers.Count -lt 3) {
                throw "Could not find valid CSV header (need at least 3 columns)"
            }
            
            Write-Output "DEBUG: Found $($lines.Count) lines, $(($headers).Count) columns"
            
            $ProductData = @()
            for ($i = $headerLineIndex + 1; $i -lt $lines.Count; $i++) {
                if (($i % 500) -eq 0) {
                    Write-Output "DEBUG: Parsed $i rows so far..."
                }
                $line = $lines[$i]
                if ($line.Trim().Length -eq 0) { continue }
                
                $values = @()
                $current = ""
                $inQuotes = $false
                for ($j = 0; $j -lt $line.Length; $j++) {
                    $char = $line[$j]
                    if ($char -eq '"') {
                        $inQuotes = -not $inQuotes
                    }
                    elseif ($char -eq ',' -and -not $inQuotes) {
                        $values += $current.Trim('"').Trim()
                        $current = ""
                    }
                    else {
                        $current += $char
                    }
                }
                $values += $current.Trim('"').Trim()
                
                $obj = [PSCustomObject]@{}
                for ($k = 0; $k -lt $headers.Count; $k++) {
                    $obj | Add-Member -NotePropertyName $headers[$k] -NotePropertyValue $values[$k]
                }
                $ProductData += $obj
            }
            
            $sw.Stop()
            Write-Output "DEBUG: CSV parsed in $($sw.ElapsedMilliseconds)ms, row count: $($ProductData.Count)"
            
            Write-Output "DEBUG: Exporting to cache..."
            $ProductData | Export-Csv -Path $ProductNamesFile -NoTypeInformation -Encoding UTF8
            Write-Output "✓ Product names downloaded and cached successfully"
            $ProductDataAvailable = $true
        }
        Else {
            Write-Output "✗ No CSV download link found on documentation page"
            Throw "Could not find CSV download link on licensing page"
        }
    }
    Catch {
        Write-Output "✗ Could not download from documentation page"
        Write-Output "Error Type: $($_.Exception.GetType().FullName)"
        Write-Output "Error Message: $($_.Exception.Message)"
        
        # Final fallback: Try to use cached version
        If (Test-Path -Path $ProductNamesFile) {
            Write-Output "DEBUG: Cache file exists at: $ProductNamesFile"
            Write-Output "DEBUG: Attempting to load cached product names..."
            Try {
                $sw = [System.Diagnostics.Stopwatch]::StartNew()
                $ProductData = Import-Csv -Path $ProductNamesFile
                $sw.Stop()
                Write-Output "DEBUG: Cache loaded in $($sw.ElapsedMilliseconds)ms"
                $rowCount = ($ProductData | Measure-Object).Count
                Write-Output "✓ Loaded $rowCount products from cache"
                $ProductDataAvailable = $true
            }
            Catch {
                Write-Output "✗ Could not load cached file"
                Write-Output "Error Type: $($_.Exception.GetType().FullName)"
                Write-Output "Error Message: $($_.Exception.Message)"
            }
        }
        Else {
            Write-Output "DEBUG: No cache file found at: $ProductNamesFile"
        }
    }
}

Write-Output "DEBUG: ProductDataAvailable = $ProductDataAvailable"
If (-not $ProductDataAvailable) {
    Write-Output "✓ Fallback: Using built-in license name mappings"
}

If ($ProductDataAvailable) {
    # If the product data file is available, use it to populate some hash tables to use to resolve SKU and service plan names
    [array]$ProductInfo = $ProductData | Sort-Object GUID -Unique
    # Create Hash table of the SKUs used in the tenant with the product display names from the Microsoft data file
    $TenantSkuHash = @{}
    ForEach ($P in $SKUs) { 
        $ProductDisplayName = $ProductInfo | Where-Object { $_.GUID -eq $P.SkuId } | `
            Select-Object -ExpandProperty Product_Display_Name
        If ($Null -eq $ProductDisplayName) {
            # Try built-in names if Microsoft data doesn't have it
            If ($BuiltInSkuNames.ContainsKey($P.SkuId)) {
                $ProductDisplayName = $BuiltInSkuNames[$P.SkuId]
            }
            Else {
                $ProductDisplayname = $P.SkuPartNumber
            }
        }
        $TenantSkuHash.Add([string]$P.SkuId, [string]$ProductDisplayName) 
    }
    # Extract service plan information and build a hash table
    [array]$ServicePlanData = $ProductData | Select-Object Service_Plan_Id, Service_Plan_Name, Service_Plans_Included_Friendly_Names | `
        Sort-Object Service_Plan_Id -Unique
    $ServicePlanHash = @{}
    ForEach ($SP in $ServicePlanData) { 
        $ServicePlanHash.Add([string]$SP.Service_Plan_Id, [string]$SP.Service_Plans_Included_Friendly_Names)
    }
}
Else {
    # If Microsoft data is not available, use built-in names
    Write-Host "Using built-in license name mappings..."
    $TenantSkuHash = @{}
    ForEach ($P in $SKUs) {
        If ($BuiltInSkuNames.ContainsKey($P.SkuId)) {
            $ProductDisplayName = $BuiltInSkuNames[$P.SkuId]
        }
        Else {
            $ProductDisplayName = $P.SkuPartNumber
        }
        $TenantSkuHash.Add([string]$P.SkuId, [string]$ProductDisplayName)
    }
}

# Generate a report about the subscriptions used in the tenant
Write-Host "Generating product subscription information..."
$SkuReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($Sku in $Skus) {
    $AvailableUnits = ($Sku.PrepaidUnits.Enabled - $Sku.ConsumedUnits)
    
    # Get friendly name from hash table, or built-in names, or fall back to SKU part number
    If ($TenantSkuHash -and $TenantSkuHash.ContainsKey($Sku.SkuId)) {
        $SkuDisplayName = $TenantSkuHash[$Sku.SkuId]
    }
    ElseIf ($BuiltInSkuNames.ContainsKey($Sku.SkuId)) {
        $SkuDisplayName = $BuiltInSkuNames[$Sku.SkuId]
    }
    Else {
        $SkuDisplayName = $Sku.SkuPartNumber
    }
    
    $DataLine = [PSCustomObject][Ordered]@{
        'License Name'    = $SkuDisplayName
        'SKU Part Number' = $Sku.SkuPartNumber
        'SkuId'           = $Sku.SkuId
        'Active'          = $Sku.PrepaidUnits.Enabled
        'Warning'         = $Sku.PrepaidUnits.Warning
        'In Use'          = $Sku.ConsumedUnits
        'Available'       = $AvailableUnits        
    }
    $SkuReport.Add($Dataline)
}

# Export CSV so monthly attachments include the data
Try {
    $SkuReport | Select-Object 'License Name', 'SKU Part Number', 'SkuId', 'Active', 'Warning', 'In Use', 'Available' | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding UTF8 -Force
    Write-Host "CSV report exported to: $CSVOutputFile"
}
Catch {
    Write-Warning "Failed to export CSV report: $($_.Exception.Message)"
}

Write-Host "Generating report..."
Write-Output "========================================"
Write-Output "RETRIEVING ORGANIZATION INFO:"
Write-Output "========================================"
Write-Output "DEBUG: Calling Get-MgOrganization (requires Organization.Read.All application permission)"
try {
    $OrgName = (Get-MgOrganization -ErrorAction Stop).DisplayName
    Write-Output "✓ Successfully retrieved organization: $OrgName"
} catch {
    Write-Output "✗ FAILED to retrieve organization"
    Write-Output "Error Type: $($_.Exception.GetType().FullName)"
    Write-Output "Error Message: $($_.Exception.Message)"
    Write-Output "Stack Trace: $($_.ScriptStackTrace)"
    Write-Output "WARNING: This requires 'Organization.Read.All' application permission"
    Write-Output "Falling back to 'Unknown Organization'"
    $OrgName = "Unknown Organization"
}
Write-Output "========================================"
# Create the HTML report. First, define the header.
$HTMLHead = "<html>
	   <style>
	   BODY{font-family: Arial; font-size: 8pt;}
	   H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	   TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	   TD{border: 1px solid #969595; padding: 5px; }
	   td.pass{background: #B7EB83;}
	   td.warn{background: #E3242B;}
	   td.fail{background: #FF2626; color: #ffffff;}
	   td.info{background: #85D4FF;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>Microsoft 365 Subscriptions and Service Plan Report</h1></p>
           <p><h2><b>For the " + $Orgname + " tenant</b></h2></p>
           <p><h3>Generated: " + $RunDate + "</h3></p></div>"

# This section highlights subscriptions that have less than 5 remaining licenses.

# First, convert the output SKU Report to HTML and then import it into an XML structure
$HTMLTable = $SkuReport | ConvertTo-Html -Fragment
[xml]$XML = $HTMLTable
# Create an attribute class to use, name it, and append to the XML table attributes
$TableClass = $XML.CreateAttribute("class")
$TableClass.Value = "AvailableUnits"
$XML.table.Attributes.Append($TableClass) | Out-Null
# Conditional formatting for the table rows. The number of available units is in table row 6, so we update td[5]
ForEach ($TableRow in $XML.table.SelectNodes("tr")) {
    # each TR becomes a member of class "tablerow"
    $TableRow.SetAttribute("class", "tablerow")
    ## If row has TD and TD[5] is 5 or less
    If (($TableRow.td) -and ([int]$TableRow.td[5] -le 5)) {
        ## tag the TD with eirher the color for "warn" or "pass" defined in the heading
        $TableRow.SelectNodes("td")[5].SetAttribute("class", "warn")
    }
    ElseIf (($TableRow.td) -and ([int]$TableRow.td[5] -gt 5)) {
        $TableRow.SelectNodes("td")[5].SetAttribute("class", "pass")
    }
}

# Wrap the output table with a div tag
$HTMLBody = [string]::Format('<div class="tablediv">{0}</div>', $XML.OuterXml)


# End stuff to output
$HTMLtail = "</body></html>"
$HTMLReport = $HTMLHead + $HTMLBody + $HTMLtail
$HTMLReport | Out-File $ReportFile  -Encoding UTF8

Write-Host "All done. Output files are" $CSVOutputFile "and" $ReportFile

# --- Check monitored SKUs for low license counts ---

$lowLicenseAlerts = @()
# Normalize monitored SKU IDs: trim whitespace and remove empty entries
$monitoredSkuIds = @($MonitorSkuId1, $MonitorSkuId2, $MonitorSkuId3, $MonitorSkuId4, $MonitorSkuId5) |
ForEach-Object { if ($_ -ne $null) { $_.ToString().Trim() } } |
Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

if ($monitoredSkuIds.Count -gt 0) {
    Write-Output "========================================"
    Write-Output "LICENSE MONITORING:"
    Write-Output "========================================"
    Write-Output "Checking $($monitoredSkuIds.Count) monitored SKU(s) for licenses below threshold of $MinimumLicenseThreshold"
    
    foreach ($skuId in $monitoredSkuIds) {
        $normalizedSkuId = $skuId.ToString().Trim()
        # Try to match by SkuId (GUID) or SkuPartNumber (friendly code), case-insensitive
        $sku = $Skus | Where-Object {
            (($_.SkuId -ne $null) -and ([string]$_.SkuId).Trim().ToLower() -eq $normalizedSkuId.ToLower()) -or
            (($_.SkuPartNumber -ne $null) -and ([string]$_.SkuPartNumber).Trim().ToLower() -eq $normalizedSkuId.ToLower())
        } | Select-Object -First 1
        if ($sku) {
            if ($normalizedSkuId.ToLower() -eq ([string]$sku.SkuPartNumber).Trim().ToLower()) {
                Write-Output "  Matched monitored SKU by part number: $normalizedSkuId -> SkuId: $($sku.SkuId)"
            }
            $availableUnits = ($sku.PrepaidUnits.Enabled - $sku.ConsumedUnits)
            
            # Get friendly name
            if ($TenantSkuHash -and $TenantSkuHash.ContainsKey($sku.SkuId)) {
                $skuDisplayName = $TenantSkuHash[$sku.SkuId]
            }
            elseif ($BuiltInSkuNames.ContainsKey($sku.SkuId)) {
                $skuDisplayName = $BuiltInSkuNames[$sku.SkuId]
            }
            else {
                $skuDisplayName = $sku.SkuPartNumber
            }
            
            Write-Output "  SKU: $skuDisplayName (ID: $skuId) - Available: $availableUnits"
            
            if ($availableUnits -lt $MinimumLicenseThreshold) {
                $alertMessage = "WARNING: '$skuDisplayName' has only $availableUnits licenses available (threshold: $MinimumLicenseThreshold)"
                Write-Output "  $alertMessage"
                $lowLicenseAlerts += [PSCustomObject]@{
                    LicenseName       = $skuDisplayName
                    SkuId             = $skuId
                    AvailableLicenses = $availableUnits
                    Threshold         = $MinimumLicenseThreshold
                }
            }
        }
        else {
            Write-Output "  SKU ID $skuId not found in tenant"
        }
    }
    Write-Output "========================================"
}

# --- SendGrid Email Sending ---
$currentDate = Get-Date
$isFirstDayOfMonth = $currentDate.Day -eq 1

Write-Output "========================================"
Write-Output "EMAIL SCHEDULE CHECK:"
Write-Output "========================================"
Write-Output "Current Date: $($currentDate.ToString('yyyy-MM-dd'))"
Write-Output "Is First Day of Month: $isFirstDayOfMonth"
Write-Output "========================================"

$attachments = @()

# Only attach full reports on first day of month
if ($isFirstDayOfMonth) {
    $filesToAttach = @($CSVOutputFile, $ReportFile)
    foreach ($f in $filesToAttach) {
        if (Test-Path $f) {
            $ext = [IO.Path]::GetExtension($f).ToLower()
            $mimeType = if ($ext -eq ".csv") {
                "text/csv"
            }
            elseif ($ext -eq ".html") {
                "text/html"
            }
            else {
                "application/octet-stream"
            }
            $attachments += @{ content = [Convert]::ToBase64String([IO.File]::ReadAllBytes($f)); filename = [IO.Path]::GetFileName($f); type = $mimeType; disposition = "attachment" }
        }
        else {
            Write-Output "Attachment missing, skipping: $f"
        }
    }
}

# Build recipients array - ensure it's always an array even with single recipient
$toRecipients = @($ToEmail.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { @{ email = $_ } })

$emailSubject = ""
$emailContent = ""
$shouldSendEmail = $false

# Determine email content based on alerts and schedule
if ($lowLicenseAlerts.Count -gt 0) {
    # Always send alert emails when licenses are low
    $shouldSendEmail = $true
    $emailSubject = "M365 License Report - LOW LICENSE ALERT"
    $alertText = "`n`n=== LOW LICENSE ALERTS ===`n"
    foreach ($alert in $lowLicenseAlerts) {
        $alertText += "- $($alert.LicenseName): $($alert.AvailableLicenses) available (threshold: $($alert.Threshold))`n"
    }
    
    if ($isFirstDayOfMonth) {
        $emailContent = "WARNING: One or more monitored licenses are below the threshold!$alertText`nSee attached full monthly reports for complete details."
    }
    else {
        $emailContent = "WARNING: One or more monitored licenses are below the threshold!$alertText`nFull reports are sent on the 1st of each month."
    }
    Write-Output "Low license alerts detected - email will be sent"
}
elseif ($isFirstDayOfMonth) {
    # Send full monthly report on 1st of month even if no alerts
    $shouldSendEmail = $true
    $emailSubject = "M365 License Report - Monthly Full Report"
    $emailContent = "This is your monthly M365 license report. See attached CSV and HTML reports for full details."
    Write-Output "First day of month - monthly full report will be sent"
}
else {
    # No alerts and not 1st of month - skip sending email
    $shouldSendEmail = $false
    Write-Output "No license alerts and not first day of month - email will be skipped"
}

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
Write-Output "========================================"

# Build the SendGrid email body with proper array formatting
# NOTE: Only include attachments field if there are actual attachments (SendGrid fails on empty array)
$emailBodyObj = @{
    personalizations = @(
        @{
            to = @($toRecipients)
        }
    )
    from             = @{ email = $FromEmail }
    subject          = $emailSubject
    content          = @(@{ type = "text/plain"; value = $emailContent })
}

# Only add attachments if there are any
if ($attachments.Count -gt 0) {
    $emailBodyObj['attachments'] = $attachments
}

$emailBody = $emailBodyObj | ConvertTo-Json -Depth 6

if ($shouldSendEmail) {
    if ($SendGridApiKey) {
        Write-Output "Attempting to send email via SendGrid..."
        try {
            $response = Invoke-RestMethod -Uri "https://api.sendgrid.com/v3/mail/send" `
                -Method Post `
                -Headers @{ "Authorization" = "Bearer $SendGridApiKey"; "Content-Type" = "application/json" } `
                -Body $emailBody
            Write-Output "SendGrid API Response: $($response | ConvertTo-Json -Compress)"
            Write-Output "✓ License report email sent successfully to $ToEmail."
        }
        catch {
            Write-Output "✗ Failed to send license report email via SendGrid"
            Write-Output "Error Message: $($_.Exception.Message)"
            if ($_.Exception.Response) {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $reader.BaseStream.Position = 0
                $reader.DiscardBufferedData()
                $responseBody = $reader.ReadToEnd()
                Write-Output "SendGrid Response Body: $responseBody"
            }
        }
    }
    else {
        Write-Output "✗ SendGrid API Key not available. Cannot send license report email."
        Write-Output "Please check KeyVault configuration: KeyVaultName='$KeyVaultName', SecretName='$SendGridSecretName'"
    }
}
else {
    Write-Output "Email sending skipped - no alerts and not first day of month"
}
Write-Output "========================================"
