$ErrorActionPreference = "Stop"

Import-Module Az.Storage
Import-Module Az.Resources

function Write-Output { param ([string] $Message) $Message }

function Grant-PIMAccess {
    param (
        [parameter(Mandatory=$true)]
        [string] $UserGroupName,

        [parameter(Mandatory=$true)]
        [string] $StorageAccountId,

        [parameter(Mandatory=$true)]
        [string] $RoleDefinitionId,

        [parameter(Mandatory=$true)]
        [int] $DurationInHours
    )

    $subscriptionId = (Get-AzContext).Subscription.SubscriptionId
    Write-Output "subscriptionId        : '$($mgGroup.Id)'"

    $fullRoleDefinitionId = "/subscriptions/$subscriptionId/providers/Microsoft.Authorization/roleDefinitions/$RoleDefinitionId"
    Write-Output "fullRoleDefinitionId  : 'fullRoleDefinitionId'"

    $expirationDuration = "PT$($DurationInHours)H"
    Write-Output "expirationDuration    : '$expirationDuration'"

    $mgGroup = Get-MgGroup -Filter "displayName eq '$UserGroupName'"
    Write-Output "mgGroup               : '$($mgGroup.Id)'"
    
    if ($null -eq $mgGroup) {
        throw "Azure AD group '$UserGroupName' not found."
    }

    $roleAssignments = Get-AzRoleAssignment -Scope $StorageAccountId -ObjectId $mgGroup.Id -RoleDefinitionId $RoleDefinitionId

    if ($roleAssignments.Count -gt 0) {
        Write-Output "Role assignment schedule request already exists for group '$UserGroupName' and role '$RoleDefinitionId' for $DurationInHours hours."
        return
    }
    
    $request = New-AzRoleAssignmentScheduleRequest -Name (New-Guid).ToString() `
                                                   -RequestType "AdminAssign" `
                                                   -PrincipalId $mgGroup.Id `
                                                   -Scope $StorageAccountId `
                                                   -RoleDefinitionId $fullRoleDefinitionId `
                                                   -ScheduleInfoStartDateTime ((Get-Date).ToUniversalTime()).AddSeconds(5) `
                                                   -ExpirationType "AfterDuration" `
                                                   -ExpirationDuration $expirationDuration `
                                                   -Justification "Temporary PIM access for group $($mgGroup.Id) to $StorageAccountId for maintenance"


    Write-Output "Successfully assigned '$RoleDefinitionId' role to group '$UserGroupName' for $DurationInHours hours."
}


function Enable-M2MaintenanceMode {
    
    param (
        [parameter(Mandatory = $true, Position = 0)]
        [ValidateSet('dev', 'test', 'prod')]
        [string] $Environment
    )

    $resourceGroupName = "b2b-ec-$Environment"
    $storageAccountName = "unib2becop$Environment"
    $fileShareName = "var"
    $tempDir = $env:TEMP
    $maintenanceFileName = ".maintenance.flag"
    $maintenanceFilePath = Join-Path $tempDir $maintenanceFileName

    Write-Output "resourceGroupName     : '$resourceGroupName'"
    Write-Output "storageAccountName    : '$storageAccountName'"
    Write-Output "fileShareName         : '$fileShareName'"
    Write-Output "tempDir               : '$tempDir'"
    Write-Output "maintenanceFileName   : '$maintenanceFileName'"
    Write-Output "maintenanceFilePath   : '$maintenanceFilePath'"
    
    Connect-MgGraph -Identity -NoWelcome
    Write-Output "Connected to Microsoft Graph"

    $azureProfile = Connect-AzAccount -Identity
    Write-Output "Connected to subscription: '$($azureProfile.Context.Subscription.Name)'"
   
    $storageAccount = Get-AzStorageAccount -ResourceGroupName $resourceGroupName -Name $storageAccountName -Verbose
    Write-Output "storageAccount        : '$($storageAccount.Id)'"

    if ($null -eq (Get-AzStorageFile -ShareName $fileShareName -Context $storageAccount.Context -Path $maintenanceFileName -ErrorAction SilentlyContinue)){
        "" > $maintenanceFilePath
        Set-AzStorageFileContent -Context $storageAccount.Context -ShareName $fileShareName -Source $maintenanceFilePath -Path $maintenanceFileName
        Write-Output "$maintenanceFileName file uploaded"
    } else {
        Write-Output "Maintence mode is already enabled"
    }

    # Reader
    Grant-PIMAccess -UserGroupName "Monsoon Developers" -StorageAccountId $storageAccount.Id -RoleDefinitionId "acdd72a7-3385-48ef-bd42-f606fba81ae7" -DurationInHours 2

    # Storage File Data SMB Share Reader
    Grant-PIMAccess -UserGroupName "Monsoon Developers" -StorageAccountId $storageAccount.Id -RoleDefinitionId "aba4ae5f-2193-4029-9191-0cb91df5e314" -DurationInHours 2

    # Storage File Data Privileged Reader
    Grant-PIMAccess -UserGroupName "Monsoon Developers" -StorageAccountId $storageAccount.Id -RoleDefinitionId "b8eda974-7b85-4f76-af95-65846b26df6d" -DurationInHours 2
}

$Environment = Get-AutomationVariable -Name Environment
Enable-M2MaintenanceMode $Environment