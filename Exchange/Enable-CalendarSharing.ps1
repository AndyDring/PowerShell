$TenantDomain = “tenant.mail.onmicrosoft.com”
$OnPremDomain = “domain.com”
$OrgGuid = (Get-OrganizationConfig).guid.tostring()
 
$EWSExternalURL = "https://mail.domain.com/EWS/Exchange.asmx"

$Credential = Get-Credential
 
$session = New-PSSession -ConfigurationName Microsoft.Exchange -Authentication Basic -ConnectionUri https://ps.outlook.com/powershell -AllowRedirection:$true -Credential $Credential
Import-PSSession $session -Prefix Cloud
 
Enable-CloudOrganizationCustomization
Set-Federationtrust -Identity 'Microsoft Federation Gateway' -RefreshMetadata:$false
Set-FederatedOrganizationIdentifier -AccountNamespace $OnPremDomain -DelegationFederationTrust 'Microsoft Federation Gateway' -Enabled:$true -DefaultDomain $null
Set-CloudFederatedOrganizationIdentifier -DefaultDomain $TenantDomain -Enabled:$true
$FederationInfo = Get-FederationInformation -DomainName $TenantDomain -BypassAdditionalDomainValidation:$true 
New-OrganizationRelationship -Name "On-premises to O365 - $OrgGuid" -TargetApplicationUri 'outlook.com' -TargetAutodiscoverEpr $FederationInfo.TargetAutodiscoverEpr.ToString() -Enabled:$true -DomainNames $TenantDomain
New-CloudOrganizationRelationship -Name "O365 to On-premises - $OrgGuid" -TargetApplicationUri "FYDIBOHF25SPDLT.$($OnPremDomain)" -TargetAutodiscoverEpr "https://autodiscover.$($OnPremDomain)/autodiscover/autodiscover.svc/WSSecurity" -Enabled:$true -DomainNames $OnPremDomain
Set-OrganizationRelationship -MailboxMoveEnabled:$true -FreeBusyAccessEnabled:$true -FreeBusyAccessLevel 'LimitedDetails' -ArchiveAccessEnabled:$true -MailTipsAccessEnabled:$true -MailTipsAccessLevel 'All' -DeliveryReportEnabled:$true -PhotosEnabled:$true -TargetOwaURL "http://outlook.com/owa/$OnPremDomain" -Identity "On-premises to O365 - $OrgGuid"
Set-CloudOrganizationRelationship -FreeBusyAccessEnabled:$true -FreeBusyAccessLevel 'LimitedDetails' -MailTipsAccessEnabled:$true -MailTipsAccessLevel 'All' -DeliveryReportEnabled:$true -PhotosEnabled:$true -Identity "O365 to On-premises - $OrgGuid”
Add-AvailabilityAddressSpace -ForestName $TenantDomain -AccessMethod 'InternalProxy' -UseServiceAccount:$true -ProxyUrl $EWSExternalURL
