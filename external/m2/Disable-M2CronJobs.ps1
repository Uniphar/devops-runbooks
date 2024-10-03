$ErrorActionPreference = "Stop"

Import-Module Az.Storage

function Write-Output { param ([string] $Message) $Message }

function Disable-M2CronJobs {
    
    param (
        [parameter(Mandatory = $true, Position = 0)]
        [ValidateSet('dev', 'test', 'prod')]
        [string] $Environment
    )

    $resourceGroupName = "b2b-ec-$Environment"
    $storageAccountName = "unib2becop$Environment"
    $fileShareName = "var"
    $tempDir = $env:TEMP
    $maintenanceFileName = ".cron_disable.flag"
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

    if ($null -eq (Get-AzStorageFile -ShareName $fileShareName -Context $storageAccount.Context -Path $maintenanceFileName -ErrorAction SilentlyContinue)){
        "" > $maintenanceFilePath
        Set-AzStorageFileContent -Context $storageAccount.Context -ShareName $fileShareName -Source $maintenanceFilePath -Path $maintenanceFileName
        Write-Output "$maintenanceFileName file created and uploaded"
    } else {
        Write-Output "Cron jobs are already disabled"
    }
}

$Environment = Get-AutomationVariable -Name Environment
Disable-M2CronJobs $Environment