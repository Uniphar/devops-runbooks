function Remove-UniRoleAssignments {
    <#

    .SYNOPSIS
    Periodically ran function that removes unwanted role assignments scoped to the current subscription.

    .DESCRIPTION
    Periodically ran function that removes role assignments to objects that are not Service Principals scoped to the current subscription. 

    .EXAMPLE
    Remove-UniRoleAssignments

    #>

    [CmdletBinding(SupportsShouldProcess)]
    Param()

    $subscriptionId = (Get-AzContext).Subscription.Id
    $scope = "/subscriptions/$subscriptionId"

	$groupNames = @("DevOps High", "DevOps Low", "Azure Contributors", "Azure Readers")

    Get-AzRoleAssignment | Where-Object Scope -eq $scope `
                         | Where-Object ObjectType -ne "ServicePrincipal" `
                         | Where-Object DisplayName -NotIn $groupNames `
                         | Remove-AzRoleAssignment -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
}

Import-Module -Name Az.Resources

Disable-AzContextAutosave -Scope Process

Connect-AzAccount -Identity

(Get-AzSubscription).Id | ForEach-Object {
    Set-AzContext -SubscriptionId $_
	Remove-UniRoleAssignments -Verbose
}