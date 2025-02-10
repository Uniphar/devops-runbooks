function Invoke-UniFrontDoorCustomDomainValidation {
    [CmdletBinding()]
    param (
        [string] $ProfileName,
        [string] $ResourceGroupName
    )

    $dnsZoneResourceGroupName = "global-dns"

    Write-Verbose "Getting custom domains for Front Door profile '$ProfileName' in resource group '$ResourceGroupName'"
    $customDomains = Get-AzFrontDoorCdnCustomDomain -ResourceGroupName $ResourceGroupName -ProfileName $ProfileName
    
    $customDomains | ForEach-Object {
        Write-Verbose "Checking custom domain '$($_.HostName)'"
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

        Write-Verbose "Updating validation token from custom domain '$($customDomain.Name)'"
        Update-AzFrontDoorCdnCustomDomainValidationToken -ResourceGroupName $ResourceGroupName -ProfileName $ProfileName -CustomDomainName $customDomain.Name

        Write-Verbose "Getting updated custom domain '$($customDomain.Name)'"
        $customDomain = Get-AzFrontDoorCdnCustomDomain -ResourceGroupName $ResourceGroupName -ProfileName $ProfileName -CustomDomainName $customDomain.Name

        Write-Verbose "Checking if _dnsauth recordset already exists for '$($customDomain.HostName)'"
        $recordSet = Get-AzDnsRecordSet -ZoneName $customDomainName -ResourceGroupName $dnsZoneResourceGroupName -Name "_dnsauth" -RecordType "TXT" -ErrorAction SilentlyContinue
    
        if ($recordSet) {
            Write-Verbose "Deleting DNS Record for '_dnsauth.$customDomainName'"
            Remove-AzDnsRecordSet -RecordSet $recordSet
        }

        Write-Verbose "Creating DNS Record for '_dnsauth.$customDomainName'"
        $records = @()
        $records += New-AzDnsRecordConfig -Value $customDomain.ValidationPropertyValidationToken
        New-AzDnsRecordSet -ZoneName $customDomainName -ResourceGroupName $dnsZoneResourceGroupName -Name "_dnsauth" -RecordType "TXT" -Ttl 3600 -DnsRecords $records
    }
}


$environment = Get-AutomationVariable -Name 'Environment'

Invoke-UniFrontDoorCustomDomainValidation -ResourceGroupName "web-$environment" -ProfileName "web-$environment-afd" -Verbose