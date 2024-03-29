param(
    [Parameter(Mandatory=$true, HelpMessage="Resource Group Name")]
    [string]$ResourceGroupName
)

Import-Module -Name Az.Resources

Connect-AzAccount -Identity

Get-AzVM -ResourceGroupName $ResourceGroupName | ForEach-Object {
    Restart-AzVM -ResourceGroupName $ResourceGroupName -Name $_.Name -ErrorAction Continue
}
