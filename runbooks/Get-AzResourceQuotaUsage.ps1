Import-Module Az.Resources

$threshold = 70
$locations = @("northeurope", "westeurope", "centralindia")
$excludeResources = @("networkwatchers")

function Write-Output { param ([string] $Message) $Message }

function Get-AzResourceQuotaUsage {
    [CmdletBinding()]
    Param(
        [parameter(Mandatory = $true, Position = 0)]
        [string] $Provider,

        [parameter(Mandatory = $true, Position = 1)]
        [string] $Location
    )

    Write-Output "Provider: $Provider"

    $subscriptionId = (Get-AzContext).Subscription.SubscriptionId
    $scope = "/subscriptions/$subscriptionId/providers/$Provider/locations/$Location"
    
    Write-Output "Getting quota limits for scope: $scope"
    $limits = Get-AzQuota -Scope $scope -ErrorAction SilentlyContinue
    if (-not $limits) {
        Write-Output "No quota limits found for scope: $scope"
        return @()
    }
    
    $limits = $limits | ForEach-Object {
        return [PSCustomObject]@{
            name  = $_.Name
            limit = $_.Limit.value     
        }
    }


    Write-Output "Getting quota usage for scope: $scope"
    $usage = Get-AzQuotaUsage -Scope $scope -ErrorAction SilentlyContinue
    if (-not $usage) {
        Write-Output "No quota usage found for scope: $scope"
        return @()
    }
    
    # $usage = $usage | Where-Object { $_.UsageValue -gt 0 }
    $quotaUsage = $usage | ForEach-Object {
        $name = $_.Name
        $currentUsage = [math]::Max($_.UsageValue, 0)
        $limit = ($limits | Where-Object { $_.name -eq $name } | Measure-Object -Property limit -Maximum | Select-Object -ExpandProperty Maximum) ?? 0              
        return [PSCustomObject]@{
            name         = $name
            currentUsage = $currentUsage
            limit        = $limit
            usagePercent = ($limit -eq 0) ? 0 : ($currentUsage * 100 / $limit)
            Type         = $Provider
            location     = $Location
        }
    }
    
    return @($quotaUsage)
}


Disable-AzContextAutosave -Scope Process

$azureProfile = Connect-AzAccount -Identity
Write-Output "Connected to subscription: '$($azureProfile.Context.Subscription.Name)'"

$resourceProviders = Get-AzResourceProvider | Where-Object { $_.RegistrationState -eq "Registered" } | Select-Object -ExpandProperty ProviderNamespace

$quotaUsage = @()

foreach ($provider in $ResourceProviders) {
    foreach ($location in $Locations) {
        Write-Output "Processing provider '$provider' in location '$location'"
        $quotaUsage += Get-AzResourceQuotaUsage -Provider $provider -Location $location -Verbose
    }
}

$resources = @($quotaUsage | Where-Object { $_.limit -gt 0 -and $_.UsagePercent -gt $threshold -and $_.Name.ToLower() -notin $excludeResources })

if ($resources) {
    Write-Output "Resources that exceed the quota usage threshold:"
    Write-Output ($resources | Format-Table | Out-String)
    Write-Error ( @{
        "_source"    = "quota"
        resources = $resources
    } | ConvertTo-Json)
}
else {
    Write-Output "No resources found that exceed the quota usage threshold"
}