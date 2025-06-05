<#
.SYNOPSIS
Restarts a specified deployment in an Azure Kubernetes Service (AKS) cluster.

.DESCRIPTION
This script restarts a specified deployment in an Azure Kubernetes Service (AKS) cluster using the Azure CLI. It requires the resource group name, cluster name, subscription name, namespace, and deployment name as parameters.

.PARAMETER rgName
The name of the resource group containing the AKS cluster.
.PARAMETER clusterName
The name of the AKS cluster.
.PARAMETER subName
The name of the Azure subscription containing the AKS cluster.
.PARAMETER namespace
The Kubernetes namespace where the deployment is located.
.PARAMETER deployName
The name of the deployment to restart.
.EXAMPLE
.\Invoke-RestartAKSDeployment.ps1 -rgName "myResourceGroup" -clusterName "myAKSCluster" -subName "mySubscription" -namespace "default" -deployName "myDeployment"
This example restarts the deployment named "myDeployment" in the "default" namespace of the AKS cluster "myAKSCluster" located in the resource group "myResourceGroup" under the subscription "mySubscription".

#>

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

$azLoginOutput = az login --identity 2>&1
if (!$?) 
{
    throw "az login failed. Details: $azLoginOutput"
}

$getCredentialsOutput = az aks get-credentials --resource-group $rgName --name $clusterName --subscription $subName 2>&1
if (!$?)
{
    throw "az aks get-credentials failed. Details: $getCredentialsOutput"
}

$invokeCommandOutput = az aks command invoke -n $clusterName -g $rgName -c "kubectl rollout restart deployment/$deployName -n $namespace" 2>&1
if (!$?)
{
    throw "az aks command invoke failed. Details: $invokeCommandOutput"
}
