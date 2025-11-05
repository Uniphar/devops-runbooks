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
    [Parameter(Mandatory = $true)] [string]$keyVaultName,
    [Parameter(Mandatory = $true)] [string]$sendGridSecretName,
    [Parameter(Mandatory = $true)] [string]$fromEmail,
    [Parameter(Mandatory = $true)] [string]$toEmail,
    [Parameter(Mandatory = $false)] [string]$monitorSkuId1 = "",
    [Parameter(Mandatory = $false)] [string]$monitorSkuId2 = "",
    [Parameter(Mandatory = $false)] [string]$monitorSkuId3 = "",
    [Parameter(Mandatory = $false)] [string]$monitorSkuId4 = "",
    [Parameter(Mandatory = $false)] [string]$monitorSkuId5 = "",
    [Parameter(Mandatory = $false)] [int]$minimumLicenseThreshold = 5
)

# Use Azure Automation temp path ($env:TEMP) when available; fallback to C:\temp for local testing.

if ($env:AUTOMATION_ACCOUNT_NAME) {
    # In Azure Automation, $env:TEMP is the correct sandbox temporary path.
    $tempPath = $env:TEMP
}
else {
    # Local testing fallback
    $tempPath = "C:\\temp"
}
if (!(Test-Path -Path $tempPath)) {
    Write-Host "Creating temporary directory: $tempPath"
    New-Item -Path $tempPath -ItemType Directory | Out-Null
}

[string]$runDate = Get-Date -format "dd-MMM-yyyy HH:mm:ss"
$csvOutputFile = Join-Path -Path $tempPath -ChildPath "Microsoft365LicenseServicePlans.csv"
$reportFile = Join-Path -Path $tempPath -ChildPath "Microsoft365LicenseServicePlans.html"

Connect-AzAccount -Identity -ErrorAction Stop
Write-Debug "✓ Connected to Azure (Az) via managed identity"

$secret = Get-AzKeyVaultSecret -VaultName $keyVaultName -Name $sendGridSecretName -AsPlainText -ErrorAction Stop
$sendGridApiKey = $secret
Write-Debug "✓ SendGrid API Key retrieved successfully from Key Vault '$keyVaultName'"
Write-Debug "✓ Secret length: $($sendGridApiKey.Length) characters"

# Connect to Microsoft Graph using the Automation Account managed identity.
Write-Host "Connecting to Microsoft Graph using Automation Managed Identity..."

## Required application permissions (grant these to the Automation Account service principal and consent):
## - Directory.Read.All (Application): read subscribed SKUs and directory information
## - Organization.Read.All (Application): read organization details used by Get-MgOrganization
# Note: when using -Identity (managed identity) do NOT pass -Scopes; passing -Scopes with -Identity causes a ParameterSet ambiguity error.

Connect-MgGraph -Identity -NoWelcome -ErrorAction Stop
Write-Debug "✓ Connected to Microsoft Graph via Managed Identity"

# Verify connection and permissions
Write-Debug "MICROSOFT GRAPH CONNECTION DEBUG:"
    $mgContext = Get-MgContext
    Write-Debug "✓ Graph Context Retrieved:"
    Write-Debug "  Account: $($mgContext.Account)"
    Write-Debug "  AppName: $($mgContext.AppName)"
    Write-Debug "  TenantId: $($mgContext.TenantId)"
    Write-Debug "  Scopes: $($mgContext.Scopes -join ', ')"
    Write-Debug "  AuthType: $($mgContext.AuthType)"
    Write-Debug "  ContextScope: $($mgContext.ContextScope)"

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
Write-Debug "========================================"
Write-Debug "RETRIEVING SUBSCRIBED SKUs:"
Write-Debug "========================================"
Write-Debug "DEBUG: Calling Get-MgSubscribedSku (requires Directory.Read.All application permission)"
try {
    [array]$skus = Get-MgSubscribedSku -ErrorAction Stop
    Write-Debug "✓ Successfully retrieved $($skus.Count) SKUs"
} catch {
    Write-Debug "✗ FAILED to retrieve SKUs - This is likely a permissions issue"
    Write-Error "Error Type: $($_.Exception.GetType().FullName)"
    Write-Error "Error Message: $($_.Exception.Message)"
    Write-Debug "Stack Trace: $($_.ScriptStackTrace)"
    Write-Debug ""
    Write-Debug "TROUBLESHOOTING:"
    Write-Debug "1. Ensure the Automation Account managed identity has 'Directory.Read.All' application permission"
    Write-Debug "2. Admin consent must be granted for the permission (not just granted via IAM)"
    Write-Debug "3. The permission must be at APPLICATION LEVEL, not delegated"
    Write-Debug "4. Grant permission via: Entra ID > App registrations > (your app) > API permissions"
    Write-Debug "5. Verify: Entra ID > Enterprise applications > (your app) > Permissions > Admin consent status = 'Granted'"
    Write-Debug ""
    throw
}
Write-Debug "========================================"


# It's used to resolve SKU and service plan code names to human-friendly values
Write-Debug "========================================"
Write-Debug "PRODUCT NAMES DOWNLOAD:"
Write-Debug "========================================"
Write-Debug "Attempting to download latest product names from Microsoft (with short timeout)..."
$productDataAvailable = $false
$productNamesFile = Join-Path -Path $tempPath -ChildPath "ProductNames.csv"

# Try to download the latest product names CSV from Microsoft (optional, non-blocking)
try {
    # Microsoft's official Product names and service plan identifiers for licensing
    # This URL is updated regularly by Microsoft
    $DirectCsvUrl = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
    
    Write-Debug "DEBUG: Starting direct download from: $DirectCsvUrl"
    Write-Debug "DEBUG: Current time: $(Get-Date -Format 'o')"
    Write-Debug "DEBUG: Timeout set to 10 seconds"
    
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Debug "DEBUG: Invoking Invoke-WebRequest..."
    $productInfoRequest = Invoke-WebRequest -Uri $DirectCsvUrl -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
    $sw.Stop()
    Write-Debug "DEBUG: Download completed in $($sw.ElapsedMilliseconds)ms"
    
    Write-Debug "DEBUG: Response received, size: $($productInfoRequest.Content.Length) bytes"
    Write-Debug "DEBUG: Parsing CSV content using streaming parser..."
    $sw.Restart()

    # Use streaming parser instead of ConvertFrom-Csv for large files
    $csvContent = $productInfoRequest.Content
    $lines = $csvContent -split "`n" | Where-Object { $_.Trim().Length -gt 0 }

    # Find the header line (skip empty lines and find first line with reasonable column count)
    $headerLineIndex = 0
    $headers = @()
    for ($h = 0; $h -lt [Math]::Min(10, $lines.Count); $h++) {
        $testHeaders = $lines[$h] -split ',' | ForEach-Object { $_.Trim('"').Trim() }
        if ($testHeaders.Count -ge 3) {
            $headers = $testHeaders
            $headerLineIndex = $h
            break
        }
    }

    if ($headers.Count -lt 3) {
        throw "Could not find valid CSV header (need at least 3 columns)"
    }

    Write-Debug "DEBUG: Found $($lines.Count) lines, $(($headers).Count) columns"

    $productData = @()
    for ($i = $headerLineIndex + 1; $i -lt $lines.Count; $i++) {
        if (($i % 500) -eq 0) {
            Write-Debug "DEBUG: Parsed $i rows so far..."
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
            $val = if ($k -lt $values.Count) { $values[$k] } else { "" }
            $obj | Add-Member -NotePropertyName $headers[$k] -NotePropertyValue $val
        }
        $productData += $obj
    }

    $sw.Stop()
    Write-Debug "DEBUG: CSV parsed in $($sw.ElapsedMilliseconds)ms, row count: $($productData.Count)"

    Write-Debug "DEBUG: Exporting to cache file: $productNamesFile"
    $sw.Restart()
    $productData | Export-Csv -Path $productNamesFile -NoTypeInformation -Encoding UTF8
    $sw.Stop()
    Write-Debug "DEBUG: Export completed in $($sw.ElapsedMilliseconds)ms"

    Write-Debug "✓ Product names downloaded and cached successfully to: $productNamesFile"
    $productDataAvailable = $true
    
}
catch {
    Write-Debug "✗ Could not download from direct URL"
    Write-Error "Error Type: $($_.Exception.GetType().FullName)"
    Write-Error "Error Message: $($_.Exception.Message)"
    Write-Error "Error Details: $($_ | Out-String)"
    
    # Fallback: Try to find the CSV link from the documentation page
    try {
    Write-Debug "DEBUG: Attempting fallback method - fetching documentation page"
    Write-Debug "DEBUG: Current time: $(Get-Date -Format 'o')"
        $LicensingPageUrl = "https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference"
        
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        Write-Debug "DEBUG: Invoking Invoke-WebRequest for documentation page..."
    $licensingPageRequest = Invoke-WebRequest -Uri $LicensingPageUrl -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
    $sw.Stop()
    Write-Debug "DEBUG: Documentation page downloaded in $($sw.ElapsedMilliseconds)ms"
        
        Write-Debug "DEBUG: Searching for CSV link in page..."
        $downloadLink = ($licensingPageRequest.Links | Where-Object { $_.href -like '*.csv' }).href
        Write-Debug "DEBUG: Found $(@($downloadLink).Count) CSV links"
        
        if ($downloadLink) {
            # Make sure the link is absolute
            if ($downloadLink -notlike "http*") {
                $downloadLink = "https://learn.microsoft.com$downloadLink"
            }
            
            Write-Debug "DEBUG: CSV link found: $downloadLink"
            Write-Debug "DEBUG: Downloading CSV from link..."
            $sw.Restart()
            $productInfoRequest = Invoke-WebRequest -Uri $downloadLink -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
            $sw.Stop()
            Write-Debug "DEBUG: Downloaded in $($sw.ElapsedMilliseconds)ms"
            
            Write-Debug "DEBUG: Parsing CSV using streaming parser..."
            $sw.Restart()
            
            # Use streaming parser instead of ConvertFrom-Csv for large files
            $csvContent = $productInfoRequest.Content
            $lines = $csvContent -split "`n" | Where-Object { $_.Trim().Length -gt 0 }
            
            # Find the header line (skip empty lines and find first line with reasonable column count)
            $headerLineIndex = 0
            $headers = @()
            for ($h = 0; $h -lt [Math]::Min(10, $lines.Count); $h++) {
                $testHeaders = $lines[$h] -split ',' | ForEach-Object { $_.Trim('"').Trim() }
                if ($testHeaders.Count -ge 3) {  # Real header should have at least 3 columns
                    $headers = $testHeaders
                    $headerLineIndex = $h
                    break
                }
            }
            
            if ($headers.Count -lt 3) {
                throw "Could not find valid CSV header (need at least 3 columns)"
            }
            
            Write-Debug "DEBUG: Found $($lines.Count) lines, $(($headers).Count) columns"
            
            $productData = @()
            for ($i = $headerLineIndex + 1; $i -lt $lines.Count; $i++) {
                if (($i % 500) -eq 0) {
                    Write-Debug "DEBUG: Parsed $i rows so far..."
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
                $productData += $obj
            }
            
            $sw.Stop()
            Write-Debug "DEBUG: CSV parsed in $($sw.ElapsedMilliseconds)ms, row count: $($productData.Count)"
            
            Write-Debug "DEBUG: Exporting to cache..."
            $productData | Export-Csv -Path $productNamesFile -NoTypeInformation -Encoding UTF8
            Write-Debug "✓ Product names downloaded and cached successfully"
            $productDataAvailable = $true
        }
        else {
            Write-Debug "✗ No CSV download link found on documentation page"
            Throw "Could not find CSV download link on licensing page"
        }
    }
    catch {
    Write-Debug "✗ Could not download from documentation page"
    Write-Error "Error Type: $($_.Exception.GetType().FullName)"
    Write-Error "Error Message: $($_.Exception.Message)"
        
        # Final fallback: Try to use cached version
        if (Test-Path -Path $productNamesFile) {
            Write-Debug "DEBUG: Cache file exists at: $productNamesFile"
            Write-Debug "DEBUG: Attempting to load cached product names..."
            try {
                $sw = [System.Diagnostics.Stopwatch]::StartNew()
                $productData = Import-Csv -Path $productNamesFile
                $sw.Stop()
                Write-Debug "DEBUG: Cache loaded in $($sw.ElapsedMilliseconds)ms"
                $rowCount = ($productData | Measure-Object).Count
                Write-Debug "✓ Loaded $rowCount products from cache"
                $productDataAvailable = $true
            }
            catch {
                Write-Debug "✗ Could not load cached file"
                Write-Error "Error Type: $($_.Exception.GetType().FullName)"
                Write-Error "Error Message: $($_.Exception.Message)"
            }
        }
        else {
            Write-Debug "DEBUG: No cache file found at: $productNamesFile"
        }
    }
}

Write-Debug "DEBUG: productDataAvailable = $productDataAvailable"
if (-not $productDataAvailable) {
    Write-Debug "✓ Fallback: Using built-in license name mappings"
}

if ($productDataAvailable) {
    # If the product data file is available, use it to populate some hash tables to use to resolve SKU and service plan names
    [array]$productInfo = $productData | Sort-Object GUID -Unique
    # Create Hash table of the SKUs used in the tenant with the product display names from the Microsoft data file
    $tenantSkuHash = @{}
    ForEach ($P in $skus) { 
        $productDisplayName = $productInfo | Where-Object { $_.GUID -eq $P.SkuId } | `
            Select-Object -ExpandProperty Product_Display_Name
        if ($Null -eq $productDisplayName) {
            # Try built-in names if Microsoft data doesn't have it
            if ($BuiltInSkuNames.ContainsKey($P.SkuId)) {
                $productDisplayName = $BuiltInSkuNames[$P.SkuId]
            }
            else {
                $productDisplayname = $P.SkuPartNumber
            }
        }
        $tenantSkuHash.Add([string]$P.SkuId, [string]$productDisplayName) 
    }
    # Extract service plan information and build a hash table
    [array]$servicePlanData = $productData | Select-Object Service_Plan_Id, Service_Plan_Name, Service_Plans_Included_Friendly_Names | `
        Sort-Object Service_Plan_Id -Unique
    $servicePlanHash = @{}
    ForEach ($SP in $servicePlanData) { 
        $servicePlanHash.Add([string]$SP.Service_Plan_Id, [string]$SP.Service_Plans_Included_Friendly_Names)
    }
}
else {
    # If Microsoft data is not available, use built-in names
    Write-Host "Using built-in license name mappings..."
    $tenantSkuHash = @{}
    ForEach ($P in $skus) {
        if ($BuiltInSkuNames.ContainsKey($P.SkuId)) {
            $productDisplayName = $BuiltInSkuNames[$P.SkuId]
        }
        else {
            $productDisplayName = $P.SkuPartNumber
        }
        $tenantSkuHash.Add([string]$P.SkuId, [string]$productDisplayName)
    }
}

# Generate a report about the subscriptions used in the tenant
    Write-Host "Generating product subscription information..."
    $skuReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($sku in $skus) {
    $availableUnits = ($sku.PrepaidUnits.Enabled - $sku.ConsumedUnits)
    
    # Get friendly name from hash table, or built-in names, or fall back to SKU part number
    if ($tenantSkuHash -and $tenantSkuHash.ContainsKey($sku.SkuId)) {
        $skuDisplayName = $tenantSkuHash[$sku.SkuId]
    }
    elseif ($BuiltInSkuNames.ContainsKey($sku.SkuId)) {
        $skuDisplayName = $BuiltInSkuNames[$sku.SkuId]
    }
    else {
        $skuDisplayName = $sku.SkuPartNumber
    }
    
    $dataLine = [PSCustomObject][Ordered]@{
        'License Name'    = $skuDisplayName
        'SKU Part Number' = $sku.SkuPartNumber
        'SkuId'           = $sku.SkuId
        'Active'          = $sku.PrepaidUnits.Enabled
        'Warning'         = $sku.PrepaidUnits.Warning
        'In Use'          = $sku.ConsumedUnits
        'Available'       = $availableUnits        
    }
    $skuReport.Add($dataLine)
}

# Export CSV so monthly attachments include the data
try {
    $skuReport | Select-Object 'License Name', 'SKU Part Number', 'SkuId', 'Active', 'Warning', 'In Use', 'Available' | Export-Csv -Path $csvOutputFile -NoTypeInformation -Encoding UTF8 -Force
    Write-Host "CSV report exported to: $csvOutputFile"
}
catch {
    Write-Warning "Failed to export CSV report: $($_.Exception.Message)"
}

Write-Host "Generating report..."
    Write-Debug "========================================"
    Write-Debug "RETRIEVING ORGANIZATION INFO:"
    Write-Debug "========================================"
    Write-Debug "DEBUG: Calling Get-MgOrganization (requires Organization.Read.All application permission)"
$orgName = "Unknown Organization"
try {
    $orgName = (Get-MgOrganization -ErrorAction Stop).DisplayName
    Write-Debug "✓ Successfully retrieved organization: $orgName"
} catch {
    Write-Debug "✗ FAILED to retrieve organization"
    Write-Error "Error Type: $($_.Exception.GetType().FullName)"
    Write-Error "Error Message: $($_.Exception.Message)"
    Write-Debug "Stack Trace: $($_.ScriptStackTrace)"
    Write-Debug "WARNING: This requires 'Organization.Read.All' application permission"
    Write-Debug "Falling back to 'Unknown Organization'"
}
# Use debug-level output for automation logs
Write-Debug "========================================"
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
           <p><h2><b>For the " + $orgName + " tenant</b></h2></p>
           <p><h3>Generated: " + $runDate + "</h3></p></div>"

# This section highlights subscriptions that have less than 5 remaining licenses.

# First, convert the output SKU Report to HTML and then import it into an XML structure
$HTMLTable = $skuReport | ConvertTo-Html -Fragment
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
    if (($TableRow.td) -and ([int]$TableRow.td[5] -le 5)) {
        ## tag the TD with either the color for "warn" or "pass" defined in the heading
        $TableRow.SelectNodes("td")[5].SetAttribute("class", "warn")
    }
    elseif (($TableRow.td) -and ([int]$TableRow.td[5] -gt 5)) {
        $TableRow.SelectNodes("td")[5].SetAttribute("class", "pass")
    }
}

# Wrap the output table with a div tag
$HTMLBody = [string]::Format('<div class="tablediv">{0}</div>', $XML.OuterXml)


# End stuff to output
$HTMLtail = "</body></html>"
$HTMLReport = $HTMLHead + $HTMLBody + $HTMLtail
$HTMLReport | Out-File $reportFile  -Encoding UTF8

Write-Host "All done. Output files are" $csvOutputFile "and" $reportFile

# --- Check monitored SKUs for low license counts ---

# --- Check monitored SKUs for low license counts ---

$lowLicenseAlerts = @()
# Normalize monitored SKU IDs: trim whitespace and remove empty entries
$monitoredSkuIds = @($monitorSkuId1, $monitorSkuId2, $monitorSkuId3, $monitorSkuId4, $monitorSkuId5) |
ForEach-Object { if ($_ -ne $null) { $_.ToString().Trim() } } |
Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

if ($monitoredSkuIds.Count -gt 0) {
    Write-Debug "========================================"
    Write-Debug "LICENSE MONITORING:"
    Write-Debug "========================================"
    Write-Debug "Checking $($monitoredSkuIds.Count) monitored SKU(s) for licenses below threshold of $minimumLicenseThreshold"
    
    foreach ($skuId in $monitoredSkuIds) {
        $normalizedSkuId = $skuId.ToString().Trim()
        # Try to match by SkuId (GUID) or SkuPartNumber (friendly code), case-insensitive
        $sku = $skus | Where-Object {
            (($_.SkuId -ne $null) -and ([string]$_.SkuId).Trim().ToLower() -eq $normalizedSkuId.ToLower()) -or
            (($_.SkuPartNumber -ne $null) -and ([string]$_.SkuPartNumber).Trim().ToLower() -eq $normalizedSkuId.ToLower())
        } | Select-Object -First 1
        if ($sku) {
            if ($normalizedSkuId.ToLower() -eq ([string]$sku.SkuPartNumber).Trim().ToLower()) {
                Write-Debug "  Matched monitored SKU by part number: $normalizedSkuId -> SkuId: $($sku.SkuId)"
            }
            $availableUnits = ($sku.PrepaidUnits.Enabled - $sku.ConsumedUnits)
            
            # Get friendly name
            if ($tenantSkuHash -and $tenantSkuHash.ContainsKey($sku.SkuId)) {
                $skuDisplayName = $tenantSkuHash[$sku.SkuId]
            }
            elseif ($BuiltInSkuNames.ContainsKey($sku.SkuId)) {
                $skuDisplayName = $BuiltInSkuNames[$sku.SkuId]
            }
            else {
                $skuDisplayName = $sku.SkuPartNumber
            }
            
            Write-Debug "  SKU: $skuDisplayName (ID: $skuId) - Available: $availableUnits"
            
            if ($availableUnits -lt $minimumLicenseThreshold) {
                $alertMessage = "WARNING: '$skuDisplayName' has only $availableUnits licenses available (threshold: $minimumLicenseThreshold)"
                Write-Debug "  $alertMessage"
                $lowLicenseAlerts += [PSCustomObject]@{
                    LicenseName       = $skuDisplayName
                    SkuId             = $sku.SkuId
                    AvailableLicenses = $availableUnits
                    Threshold         = $minimumLicenseThreshold
                }
            }
        }
        else {
            Write-Debug "  SKU ID $skuId not found in tenant"
        }
    }
    Write-Debug "========================================"
}

# --- SendGrid Email Sending ---
$currentDate = Get-Date
$isFirstDayOfMonth = $currentDate.Day -eq 1

Write-Debug "========================================"
Write-Debug "EMAIL SCHEDULE CHECK:"
Write-Debug "========================================"
Write-Debug "Current Date: $($currentDate.ToString('yyyy-MM-dd'))"
Write-Debug "Is First Day of Month: $isFirstDayOfMonth"
Write-Debug "========================================"

$attachments = @()

# Only attach full reports on first day of month
if ($isFirstDayOfMonth) {
    $filesToAttach = @($csvOutputFile, $reportFile)
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
            Write-Debug "Attachment missing, skipping: $f"
        }
    }
}

# Build recipients array - ensure it's always an array even with single recipient
$toRecipients = @($toEmail.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { @{ email = $_ } })

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
    Write-Debug "Low license alerts detected - email will be sent"
}
elseif ($isFirstDayOfMonth) {
    # Send full monthly report on 1st of month even if no alerts
    $shouldSendEmail = $true
    $emailSubject = "M365 License Report - Monthly Full Report"
    $emailContent = "This is your monthly M365 license report. See attached CSV and HTML reports for full details."
    Write-Debug "First day of month - monthly full report will be sent"
}
else {
    # No alerts and not 1st of month - skip sending email
    $shouldSendEmail = $false
    Write-Debug "No license alerts and not first day of month - email will be skipped"
}

Write-Debug "========================================"
Write-Debug "EMAIL SENDING DEBUG INFORMATION:"
Write-Debug "========================================"
Write-Debug "SendGrid API Key present: $(if ($sendGridApiKey) { 'YES (length: ' + $sendGridApiKey.Length + ')' } else { 'NO' })"
Write-Debug "From Email: $fromEmail"
Write-Debug "To Email: $toEmail"
Write-Debug "Subject: $emailSubject"
Write-Debug "Number of attachments: $($attachments.Count)"
Write-Debug "Attachment files:"
foreach ($att in $attachments) {
    Write-Debug "  - $($att.filename) (size: $($att.content.Length) base64 chars)"
}
Write-Debug "========================================"

# Build the SendGrid email body with proper array formatting
# NOTE: Only include attachments field if there are actual attachments (SendGrid fails on empty array)
$emailBodyObj = @{
    personalizations = @(
        @{
            to = @($toRecipients)
        }
    )
    from             = @{ email = $fromEmail }
    subject          = $emailSubject
    content          = @(@{ type = "text/plain"; value = $emailContent })
}

# Only add attachments if there are any
if ($attachments.Count -gt 0) {
    $emailBodyObj['attachments'] = $attachments
}

$emailBody = $emailBodyObj | ConvertTo-Json -Depth 6

if ($shouldSendEmail) {
    if ($sendGridApiKey) {
    Write-Debug "Attempting to send email via SendGrid..."
        try {
            $response = Invoke-RestMethod -Uri "https://api.sendgrid.com/v3/mail/send" `
                -Method Post `
                -Headers @{ "Authorization" = "Bearer $sendGridApiKey"; "Content-Type" = "application/json" } `
                -Body $emailBody
            Write-Debug "SendGrid API Response: $($response | ConvertTo-Json -Compress)"
            Write-Debug "✓ License report email sent successfully to $toEmail."
        }
        catch {
            Write-Debug "✗ Failed to send license report email via SendGrid"
            Write-Error "Error Message: $($_.Exception.Message)"
            if ($_.Exception.Response) {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $reader.BaseStream.Position = 0
                $reader.DiscardBufferedData()
                $responseBody = $reader.ReadToEnd()
                Write-Debug "SendGrid Response Body: $responseBody"
            }
        }
    }
    else {
    Write-Debug "✗ SendGrid API Key not available. Cannot send license report email."
    Write-Debug "Please check KeyVault configuration: KeyVaultName='$keyVaultName', SecretName='$sendGridSecretName'"
    }
}
else {
    Write-Debug "Email sending skipped - no alerts and not first day of month"
}
Write-Debug "========================================"
