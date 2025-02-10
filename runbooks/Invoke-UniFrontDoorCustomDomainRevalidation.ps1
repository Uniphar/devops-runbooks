function Write-Output { param ([string] $Message) $Message }

function Invoke-UniFrontDoorCustomDomainRevalidation {
    [CmdletBinding()]
    param (
        [string] $ProfileName,
        [string] $ResourceGroupName
    )

    $dnsZoneResourceGroupName = "global-dns"

    Write-Output "Getting custom domains for Front Door profile '$ProfileName' in resource group '$ResourceGroupName'"
    $customDomains = Get-AzFrontDoorCdnCustomDomain -ResourceGroupName $ResourceGroupName -ProfileName $ProfileName
    
    $customDomains | ForEach-Object {
        Write-Output "Checking custom domain '$($_.HostName)'"
        $dnsZone = Get-AzDnsZone -ResourceGroupName $dnsZoneResourceGroupName -Name $_.HostName -ErrorAction SilentlyContinue
        $isApex = $dnsZone ? $true : $false
        return @{
            Name = $_.Name
            HostName = $_.HostName
            IsApex = $isApex
            HasManageCertificate = $_.TlsSetting.CertificateType -eq "ManagedCertificate"
            DomainValidationState = $_.DomainValidationState
        }
    } | Where-Object { $_.IsApex -and $_.HasManageCertificate -and $_.DomainValidationState -eq "PendingRevalidation" } | ForEach-Object {
        $customDomain = $_

        Write-Output "Updating validation token from custom domain '$($customDomain.Name)'"
        Update-AzFrontDoorCdnCustomDomainValidationToken -ResourceGroupName $ResourceGroupName -ProfileName $ProfileName -CustomDomainName $customDomain.Name

        Write-Output "Getting updated custom domain '$($customDomain.Name)'"
        $customDomain = Get-AzFrontDoorCdnCustomDomain -ResourceGroupName $ResourceGroupName -ProfileName $ProfileName -CustomDomainName $customDomain.Name

        Write-Output "Checking if _dnsauth recordset already exists for '$($customDomain.HostName)'"
        $recordSet = Get-AzDnsRecordSet -ZoneName $($customDomain.HostName) -ResourceGroupName $dnsZoneResourceGroupName -Name "_dnsauth" -RecordType "TXT" -ErrorAction SilentlyContinue
    
        if ($recordSet) {
            Write-Output "Deleting DNS Record for '_dnsauth.$($customDomain.HostName)'"
            Remove-AzDnsRecordSet -RecordSet $recordSet
        }

        Write-Output "Creating DNS Record for '_dnsauth.$($customDomain.HostName)'"
        $records = @()
        $records += New-AzDnsRecordConfig -Value $customDomain.ValidationPropertyValidationToken
        New-AzDnsRecordSet -ZoneName $($customDomain.HostName) -ResourceGroupName $dnsZoneResourceGroupName -Name "_dnsauth" -RecordType "TXT" -Ttl 3600 -DnsRecords $records
    }
}

Disable-AzContextAutosave -Scope Process
$azureProfile = Connect-AzAccount -Identity
Write-Output "Connected to subscription: '$($azureProfile.Context.Subscription.Name)'"

$environment = Get-AutomationVariable -Name 'Environment'

Invoke-UniFrontDoorCustomDomainRevalidation -ResourceGroupName "web-$environment" -ProfileName "web-$environment-afd" -Verbose