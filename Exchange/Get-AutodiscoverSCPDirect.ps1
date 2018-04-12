
$Servers = Get-ClientAccessServer | Get-ExchangeServer |Where-Object{$_.AdminDisplayVersion -match "Version 15"} |Select-Object Name
Foreach ($Server in $Servers) {
    $Name = $Server.Name
    Get-ADObject "CN=$Name,CN=Autodiscover,CN=Protocols,CN=$Name,CN=Servers,CN=Exchange Administrative Group (FYDIBOHF23SPDLT),CN=Administrative Groups,CN=PHU,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=HomeOffice,DC=Local" -Properties keywords,serviceBindingInformation |Format-List Name,keywords,serviceBindingInformation
