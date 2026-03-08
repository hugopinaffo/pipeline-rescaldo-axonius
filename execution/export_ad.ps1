param([string]$OutputPath)
Import-Module ActiveDirectory -ErrorAction Stop

Get-ADComputer -Filter * -Properties OperatingSystem,lastLogon,lastLogonTimestamp,Created,Enabled,DistinguishedName |
Select-Object Name,OperatingSystem,
    @{N='lastLogon';E={[datetime]::FromFileTime($_.lastLogon)}},
    @{N='lastLogonTimestamp';E={[datetime]::FromFileTime($_.lastLogonTimestamp)}},
    Created,Enabled,DistinguishedName |
Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
