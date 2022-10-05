Param(
    [parameter(Mandatory = $true, Position = 0)]
    [string] $ResourceGroupName,

    [parameter(Mandatory = $true, Position = 1)]
    [string] $VNetGwName,

    [parameter(Mandatory = $true, Position = 2)]
    [string] $RouteTableName
)

function Update-UniHubRoutes {
<#

.SYNOPSIS
Periodically ran function that updates hub route table.

.DESCRIPTION
Periodically ran function that updates hub route table to include all routes to on prem learned by the Virtual Network Gateway through BGP.
The Hub's route table defines the Azure Firewall as the next hop for all routes to on prem.

.EXAMPLE
Update-UniHubRoutes $ResourceGroupName $VNetGwName $RouteTableName

#>

    [CmdletBinding(SupportsShouldProcess)]
    Param(
        [parameter(Mandatory = $true, Position = 0)]
        [string] $ResourceGroupName,

        [parameter(Mandatory = $true, Position = 1)]
        [string] $VNetGwName,

        [parameter(Mandatory = $true, Position = 2)]
        [string] $RouteTableName
    )

    $routeTable = Get-AzRouteTable -ResourceGroupName $ResourceGroupName -Name $RouteTableName
    $learnedRoutes = (Get-AzVirtualNetworkGatewayLearnedRoute -ResourceGroupName $ResourceGroupName -VirtualNetworkGatewayname $VNetGwName | Where-Object Origin -eq "EBgp" | Where-Object Network -ne '169.254.21.0/30')
    $firewallIp = (Get-AzFirewall -ResourceGroupName $ResourceGroupName).IpConfigurations.PrivateIpAddress

    foreach ($route in $learnedRoutes) {
        if ($null -eq (Get-AzRouteConfig -RouteTable $routeTable -Name "EBgp-$($route.Network.Replace('/','-'))_" -ErrorAction SilentlyContinue)) {
            Add-AzRouteConfig -RouteTable $routeTable `
                              -Name "EBgp-$($route.Network.Replace('/','-'))_" `
                              -AddressPrefix $route.Network `
                              -NextHopType "VirtualAppliance" `
                              -NextHopIpAddress $firewallIp
        }
    }

    Set-AzRouteTable -RouteTable $routeTable
}

Disable-AzContextAutosave -Scope Process
Connect-AzAccount -Identity
Set-AzContext -SubscriptionName uniphar-platform

Update-UniHubRoutes $ResourceGroupName $VNetGwName $RouteTableName