$ErrorActionPreference = "Stop"

Import-Module Az.Storage

function Write-Output { param ([string] $Message) $Message }

function Enable-M2CronJobs {
    
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

    if ($null -ne (Get-AzStorageFile -ShareName $fileShareName -Context $storageAccount.Context -Path $maintenanceFileName -ErrorAction SilentlyContinue)){
        Remove-AzStorageFile -Context $storageAccount.Context  -ShareName $fileShareName -Path $maintenanceFileName
        Write-Output "$maintenanceFileName file deleted"
    } else {
        Write-Output "Cron jobs are already enabled"
    }
}

$Environment = Get-AutomationVariable -Name Environment
Enable-M2CronJobs $Environment