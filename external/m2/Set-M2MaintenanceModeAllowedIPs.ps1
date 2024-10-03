[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Comma separated list of allowed IPs")]
    [string]$AllowedIPs
)

$ErrorActionPreference = "Stop"

Import-Module Az.Storage
Import-Module Az.Resources

function Write-Output { param ([string] $Message) $Message }

function Set-M2MaintenanceModeAllowedIPs {
    
    param (
        [parameter(Mandatory = $true, Position = 0)]
        [ValidateSet('dev', 'test', 'prod')]
        [string] $Environment,

        [parameter(Mandatory = $true, Position = 1)]
        [string] $AllowedIPs
    )

    $resourceGroupName = "b2b-ec-$Environment"
    $storageAccountName = "unib2becop$Environment"
    $fileShareName = "var"
    $tempDir = $env:TEMP
    $maintenanceFileName = ".maintenance.ip"
    $maintenanceFilePath = Join-Path $tempDir $maintenanceFileName

    Write-Output "resourceGroupName     : '$resourceGroupName'"
    Write-Output "storageAccountName    : '$storageAccountName'"
    Write-Output "fileShareName         : '$fileShareName'"
    Write-Output "tempDir               : '$tempDir'"
    Write-Output "maintenanceFileName   : '$maintenanceFileName'"
    Write-Output "maintenanceFilePath   : '$maintenanceFilePath'"
    
    $azureProfile = Connect-AzAccount -Identity
    Write-Output "Connected to subscription: '$($azureProfile.Context.Subscription.Name)'"
   
    $storageAccount = Get-AzStorageAccount -ResourceGroupName $resourceGroupName -Name $storageAccountName -Verbose
    Write-Output "storageAccount        : '$($storageAccount.Id)'"

    $AllowedIPs > $maintenanceFilePath
    Set-AzStorageFileContent -Context $storageAccount.Context -ShareName $fileShareName -Source $maintenanceFilePath -Path $maintenanceFileName -Force
    Write-Output "$maintenanceFileName file uploaded"
}

$Environment = Get-AutomationVariable -Name Environment
Set-M2MaintenanceModeAllowedIPs $Environment $AllowedIPs