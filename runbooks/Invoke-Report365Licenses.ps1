# To be checked:
# How to mount the storage account 
# Setup send email in the end of the script 

[string]$RunDate = Get-Date -format "dd-MMM-yyyy HH:mm:ss"
$CSVOutputFile = "\\unistorageprodkrb.file.core.windows.net\files-it\Automations\Licenses\Microsoft365LicenseServicePlans.csv"
$ReportFile = "\\unistorageprodkrb.file.core.windows.net\files-it\Automations\Licenses\Microsoft365LicenseServicePlans.html"
$SkuExceptionsFile = "\\unistorageprodkrb.file.core.windows.net\files-it\Automations\Licenses\SkuExceptions.csv"

# Connect to the Graph and get information about the subscriptions in the tenant
Connect-MgGraph -Identity -NoWelcome

Disable-AzContextAutosave -Scope Process
$context = (Connect-AzAccount -Identity).context
Set-AzContext -SubscriptionName $context.Subscription -DefaultProfile $context

# Get the basic information about tenant subscriptions
[array]$Skus = Get-MgSubscribedSku

# Load SKU exceptions
If (Test-Path -Path $SkuExceptionsFile) {
    $SkuExceptions = Import-Csv $SkuExceptionsFile | Select-Object -ExpandProperty SkuId
} Else {
    Write-Host "No SKU exceptions file available"
    $SkuExceptions = @()
}

# Filter out the SKUs that are in the exceptions list
$Skus = $Skus | Where-Object { $SkuExceptions -notcontains $_.SkuId }

# The $ProductInfoDataFile variable points to the CSV file downloaded from Microsoft from
# https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
# It's used to resolve SKU and service plan code names to human-friendly values
Write-Output "Loading product data..."
$ProductInfoDataFile = "\\unistorageprodkrb.file.core.windows.net\files-it\Automations\Licenses\Product names and service plan identifiers for licensing.csv"
If (!(Test-Path -Path $ProductInfoDataFile)) {
    Write-Host "No product information data file available - product and service plan names will not be resolved"
    $ProductData = $false
} Else {
    $ProductData = $true
}

If ($ProductData) {
# If the product data file is available, use it to populate some hash tables to use to resolve SKU and service plan names
    [array]$ProductData = Import-CSV $ProductInfoDataFile
    [array]$ProductInfo = $ProductData | Sort-Object GUID -Unique
    # Create Hash table of the SKUs used in the tenant with the product display names from the Microsoft data file
    $TenantSkuHash = @{}
        ForEach ($P in $SKUs) { 
            $ProductDisplayName = $ProductInfo | Where-Object {$_.GUID -eq $P.SkuId} | `
                Select-Object -ExpandProperty Product_Display_Name
            If ($Null -eq $ProductDisplayName) {
                $ProductDisplayname = $P.SkuPartNumber
            }
            $TenantSkuHash.Add([string]$P.SkuId, [string]$ProductDisplayName) 
        }
# Extract service plan information and build a hash table
    [array]$ServicePlanData = $ProductData | Select-Object Service_Plan_Id, Service_Plan_Name, Service_Plans_Included_Friendly_Names | `
        Sort-Object Service_Plan_Id -Unique
    $ServicePlanHash = @{}
    ForEach ($SP in $ServicePlanData) { 
        $ServicePlanHash.Add([string]$SP.Service_Plan_Id,[string]$SP.Service_Plans_Included_Friendly_Names)
    }
}

# Generate a report about the subscriptions used in the tenant
Write-Host "Generating product subscription information..."
$SkuReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($Sku in $Skus) {
    $AvailableUnits = ($Sku.PrepaidUnits.Enabled - $Sku.ConsumedUnits)
    If ($ProductData) {
        $SkuDisplayName = $TenantSkuHash[$Sku.SkuId]
    } Else {
        $SkuDisplayName = $Sku.SkuPartNumber
    }
    $DataLine = [PSCustomObject][Ordered]@{
        'Sku Part Number'   = $SkuDisplayName
        'SkuId'             = $Sku.SkuId
        'Active'            = $Sku.PrepaidUnits.Enabled
        'Warning'           = $Sku.PrepaidUnits.Warning
        'In Use'            = $Sku.ConsumedUnits
        'Available'         = $AvailableUnits        
    }
    $SkuReport.Add($Dataline)
}

Write-Host "Generating report..."
$OrgName  = (Get-MgOrganization).DisplayName
# Create the HTML report. First, define the header.
$HTMLHead="<html>
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
    $TableRow.SetAttribute("class","tablerow")
    ## If row has TD and TD[5] is 5 or less
    If (($TableRow.td) -and ([int]$TableRow.td[5] -le 1))  {
        ## tag the TD with eirher the color for "warn" or "pass" defined in the heading
        $TableRow.SelectNodes("td")[5].SetAttribute("class","warn")
    } ElseIf (($TableRow.td) -and ([int]$TableRow.td[5] -gt 5)) {
        $TableRow.SelectNodes("td")[5].SetAttribute("class","pass")
    }
}

# Wrap the output table with a div tag
$HTMLBody = [string]::Format('<div class="tablediv">{0}</div>',$XML.OuterXml)


# End stuff to output

$HTMLReport = $HTMLHead + $HTMLBody + $HTMLSkuSeparator + $HTMLSkuOutput + $HTMLtail
$HTMLReport | Out-File $ReportFile  -Encoding UTF8

Write-Host "All done. Output files are" $CSVOutputFile "and" $ReportFile