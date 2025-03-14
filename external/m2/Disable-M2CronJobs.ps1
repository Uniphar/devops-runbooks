$ErrorActionPreference = "Stop"

Import-Module Az.Storage

function Write-Output { param ([string] $Message) $Message }

function Disable-M2CronJobs {
    
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

    if ($null -eq (Get-AzStorageFile -ShareName $fileShareName -Context $storageAccount.Context -Path $maintenanceFileName -ErrorAction SilentlyContinue)){
        "" > $maintenanceFilePath
        Set-AzStorageFileContent -Context $storageAccount.Context -ShareName $fileShareName -Source $maintenanceFilePath -Path $maintenanceFileName
        Write-Output "$maintenanceFileName file created and uploaded"
    } else {
        Write-Output "Cron jobs are already disabled"
    }
}

$StorageAccountName = Get-AutomationVariable -Name 'M2_OperationsStorageAccountName'
$ResourceGroupName = Get-AutomationVariable -Name 'M2_ResourceGroupName'
Disable-M2CronJobs $StorageAccountName $ResourceGroupName