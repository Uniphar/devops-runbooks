$ErrorActionPreference = "Stop"

Import-Module Az.Storage

function Write-Output { param ([string] $Message) $Message }

function Enable-M2CronJobs {
    
    param (
        [parameter(Mandatory = $true, Position = 0)]
        [string] $StorageAccountName,

        [parameter(Mandatory = $true, Position = 1)]
        [string] $ResourceGroupName
    )

    $fileShareName = "var"
    $tempDir = $env:TEMP
    $maintenanceFileName = ".cron_disable.flag"
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

    if ($null -ne (Get-AzStorageFile -ShareName $fileShareName -Context $storageAccount.Context -Path $maintenanceFileName -ErrorAction SilentlyContinue)){
        Remove-AzStorageFile -Context $storageAccount.Context  -ShareName $fileShareName -Path $maintenanceFileName
        Write-Output "$maintenanceFileName file deleted"
    } else {
        Write-Output "Cron jobs are already enabled"
    }
}

$StorageAccountName = Get-AutomationVariable -Name 'M2_OperationsStorageAccountName'
$ResourceGroupName = Get-AutomationVariable -Name 'M2_ResourceGroupName'
Enable-M2CronJobs $StorageAccountName $ResourceGroupName