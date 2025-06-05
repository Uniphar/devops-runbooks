[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string] $rgName,

    [Parameter(Mandatory = $true)]
    [string] $clusterName,

    [Parameter(Mandatory = $true)]
    [string] $subName,

    [Parameter(Mandatory = $true)]
    [string] $namespace,
    
    [Parameter(Mandatory = $true)]
    [string] $deployName
)

az login --identity
if (!$?) 
{
    throw "az login failed"
}

az aks get-credentials --resource-group $rgName --name $clusterName --subscription $subName
if (!$?)
{
    throw "az aks get-credentials failed"
}

az aks command invoke -n $clusterName -g $rgName -c "kubectl rollout restart deployment/$deployName -n $namespace"
if (!$?)
{
    throw "az aks command invoke failed"
}
