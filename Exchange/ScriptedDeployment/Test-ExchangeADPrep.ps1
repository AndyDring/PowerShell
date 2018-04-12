#Requires -Module ActiveDirectory

<#
.NOTES
Written by Andy Dring
April 2018
Andy@AndyDring.IT


.SYNOPSIS
Script to extract various Exchange values from AD, to verify that they're at the correct level for an Exchange install. 
Also launches MS TechNet article that lists the expected values for each version of Exchange/UR/CU

.DESCRIPTION
Queries for:
RangeUpper from Schema Partition
(CN=ms-Exch-Schema-Version-Pt,CN=Schema,CN=Configuration,$DomainComponents)

Exchange Org Version from Configuration Partition
(CN=<EXCHANGE ORG>,CN=Microsoft Exchange,CN=Services,CN=Configuration,$DomainComponents)

Exchange Product ID from Configuration Partition
(CN=<EXCHANGE ORG>,CN=Microsoft Exchange,CN=Services,CN=Configuration,$DomainComponents)

Object Version from MESO Container
(CN=Microsoft Exchange System Objects,$DomainComponents)

.EXAMPLE
D:\Build\Test-ExchangeADPrep.ps1


.NOTES
Replace Active Directory LDAP Domain Components ($DomainComponents) with relevant FQDN details

Replace $ExchangeOrgName = "OrgName" with Exchange Organisation Name
#>

$ExchangeProperties = @()
$ExchangeOrgName = "OrgName"
$DomainComponents = "DC=domain,DC=local"

$ExchangeProperty = New-Object System.Object

$ExchangeProperty | Add-Member -MemberType NoteProperty -Name "Name" -Value "RangeUpper"
$ExchangeProperty | Add-Member -MemberType NoteProperty -Name "Value" -Value (Get-ADObject -Identity "CN=ms-Exch-Schema-Version-Pt,CN=Schema,CN=Configuration,$DomainComponents" -prop rangeUpper).rangeUpper
$ExchangeProperties += $ExchangeProperty


$ExchangeProperty = New-Object System.Object

$ExchangeProperty | Add-Member -MemberType NoteProperty -Name "Name" -Value "Organisation Object Version"
$ExchangeProperty | Add-Member -MemberType NoteProperty -Name "Value" -Value (Get-ADObject -Identity "CN=$ExchangeOrgName,CN=Microsoft Exchange,CN=Services,CN=Configuration,$DomainComponents" -prop objectVersion).objectVersion
$ExchangeProperties += $ExchangeProperty


$ExchangeProperty = New-Object System.Object

$ExchangeProperty | Add-Member -MemberType NoteProperty -Name "Name" -Value "Exchange Product ID"
$ExchangeProperty | Add-Member -MemberType NoteProperty -Name "Value" -Value (Get-ADObject -Identity "CN=$ExchangeOrgName,CN=Microsoft Exchange,CN=Services,CN=Configuration,$DomainComponents" -prop msExchProductID).msExchProductID
$ExchangeProperties += $ExchangeProperty


$ExchangeProperty = New-Object System.Object

$ExchangeProperty | Add-Member -MemberType NoteProperty -Name "Name" -Value "MESO Object Version"
$ExchangeProperty | Add-Member -MemberType NoteProperty -Name "Value" -Value (Get-ADObject -Identity "CN=Microsoft Exchange System Objects,$DomainComponents" -prop objectVersion).objectVersion
$ExchangeProperties += $ExchangeProperty


"Current Exchange Attribute Values from AD:"
$ExchangeProperties |Format-Table -AutoSize

Write-Host "Please compare to the Exchange Schema Values TechNet Article"
Start-Process "https://technet.microsoft.com/en-us/library/bb125224%28v=exchg.150%29.aspx?f=255&MSPPError=-2147217396#ADversions"
