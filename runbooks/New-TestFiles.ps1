Param(
    [parameter(Mandatory = $true, Position = 0)]
    [string[]] $FileIds,

    [parameter(Mandatory = $true, Position = 1)]
    [string] $FileShareName,

    [parameter(Mandatory = $true, Position = 2)]
    [string] $StorageAccountName,

    [parameter(Mandatory = $true, Position = 3)]
    [string] $StorageAccountRgName
)

Import-Module Az.Storage

function New-TestFiles {
<#

.SYNOPSIS
Generates .back files to an Azure File Share for testing purposes.

.DESCRIPTION
Generates .back files to an Azure File Share for testing purposes.
To be ran in an automation runbook, this function takes in the following parameters: 
[string[]] FileIds, [string] FileShareName, [string] StorageAccountName, [string] StorageAccountRGName

To run this function locally, provide the following parameters:
[string] FileIds - File identifiers (beginning of file names);
[Microsoft.WindowsAzure.Commands.Common.Storage.ResourceModel.AzureStorageFileShare] FileShare;

.PARAMETER FileIds
[string[]] 
File identifiers. Each string in the array represents the beginning of a file name. 
The number of files created is equal to the number of strings in the array.

.PARAMETER FileShare
[Microsoft.WindowsAzure.Commands.Common.Storage.ResourceModel.AzureStorageFileShare]
Fileshare in which the files are to be created.

.EXAMPLE
In a runbook:
New-TestFiles $FileIds $FileShareName $StorageAccountName $StorageAccountRgName

To run locally:
$fs = Get-AzStorageAccount -Name <storage account name> -ResourceGroupName <RG name> | Get-AzStorageShare -Name <file share name>
$fs | New-TestFiles @('FILE-1', 'FILE-2', 'FOO-1')

#>

    [CmdletBinding(DefaultParameterSetName = 'Runbook')]
    Param(
        [parameter(Mandatory = $true, Position = 0)]
        [string[]] $FileIds,

        [parameter(Mandatory = $true, ParameterSetName = 'Local', ValuefromPipeline = $True, Position = 1)]
        [Microsoft.WindowsAzure.Commands.Common.Storage.ResourceModel.AzureStorageFileShare] $FileShare,

        [parameter(Mandatory = $true, ParameterSetName = 'Runbook', Position = 1)]
        [string] $FileShareName,

        [parameter(Mandatory = $true, ParameterSetName = 'Runbook', Position = 2)]
        [string] $StorageAccountName,

        [parameter(Mandatory = $true, ParameterSetName = 'Runbook', Position = 3)]
        [string] $StorageAccountRgName
    )

    if(!$FileShare) {
        Disable-AzContextAutosave -Scope Process
        $AzureContext = (Connect-AzAccount -Identity).context
        $AzureContext = Set-AzContext -SubscriptionName $AzureContext.Subscription -DefaultProfile $AzureContext
        $storageContext = (Get-AzStorageAccount -ResourceGroupName $StorageAccountRgName -Name $StorageAccountName).context
    }

    foreach($id in $FileIds) {
        $guidPart = (New-Guid).ToString().Replace('-', '').Substring(0,12)
        $fileName = "$id-$guidPart.back"
        $file = New-Item -Name $fileName -ItemType File -Path $env:TEMP
        if(!$FileShare) {
            Set-AzStorageFileContent -ShareName $FileShareName -Context $storageContext -Source $file
        }
        else {
            Set-AzStorageFileContent -Share $FileShare.CloudFileShare -Source $file
        }
    }
}

New-TestFiles $FileIds $FileShareName $StorageAccountName $StorageAccountRgName