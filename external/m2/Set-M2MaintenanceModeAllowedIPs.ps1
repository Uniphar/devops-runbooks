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
        [string] $StorageAccountName,

        [parameter(Mandatory = $true, Position = 1)]
        [string] $ResourceGroupName,

        [parameter(Mandatory = $true, Position = 2)]
        [string] $AllowedIPs
    )

    $fileShareName = "var"
    $tempDir = $env:TEMP
    $maintenanceFileName = ".maintenance.ip"
    $maintenanceFilePath = Join-Path $tempDir $maintenanceFileName

    Write-Output "resourceGroupName     : '$ResourceGroupName'"
    Write-Output "storageAccountName    : '$StorageAccountName'"
    Write-Output "fileShareName         : '$fileShareName'"
    Write-Output "tempDir               : '$tempDir'"
    Write-Output "maintenanceFileName   : '$maintenanceFileName'"
    Write-Output "maintenanceFilePath   : '$maintenanceFilePath'"
    
    $azureProfile = Connect-AzAccount -Identity
    Write-Output "Connected to subscription: '$($azureProfile.Context.Subscription.Name)'"
   
    $storageAccount = Get-AzStorageAccount -ResourceGroupName $ResourceGroupName -Name $StorageAccountName -Verbose
    Write-Output "storageAccount        : '$($storageAccount.Id)'"

    $AllowedIPs > $maintenanceFilePath
    Set-AzStorageFileContent -Context $storageAccount.Context -ShareName $fileShareName -Source $maintenanceFilePath -Path $maintenanceFileName -Force
    Write-Output "$maintenanceFileName file uploaded"
}

$StorageAccountName = Get-AutomationVariable -Name 'M2_OperationsStorageAccountName'
$ResourceGroupName = Get-AutomationVariable -Name 'M2_ResourceGroupName'
Set-M2MaintenanceModeAllowedIPs $StorageAccountName $ResourceGroupName $AllowedIPs