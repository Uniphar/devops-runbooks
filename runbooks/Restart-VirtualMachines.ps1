<#
    .SYNOPSIS
    Restarts all virtual machines in a resource group.

    .DESCRIPTION
    Restarts all virtual machines in a resource group.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Resource Group Name")]
    [string]$ResourceGroupName
)

Import-Module -Name Az.Resources

Connect-AzAccount -Identity

Get-AzVM -ResourceGroupName $ResourceGroupName | ForEach-Object {
    Restart-AzVM -ResourceGroupName $ResourceGroupName -Name $_.Name -NoWait -ErrorAction Continue
}
