#Requires -Module ActiveDirectory

<#
.NOTES
Written by Andy Dring
April 2018
Andy@AndyDring.IT

.SYNOPSIS
Script to test connectivity from new Exchange Hybrid server to Exchange, DCs and EOL endpoints

.DESCRIPTION
Tests included:
Ping to all other new servers
Ping to existing Exchange servers
TCP connectivity to existing Exchange servers, on ports 80,443,25,26,135
Ping to all DCs in same site and one DC in each site with a shared Site-Link Connector
TCP connectivity to the above DCs on ports 389, 3268, 88, 135
TCP connectivity to all configured DNS servers on port 53
TCP connectivity to various Exchange Online endpoints, as per 
https://support.office.com/en-gb/article/Office-365-URLs-and-IP-address-ranges-8548a211-3fe7-47cb-abb1-355ea5aa88a2

Required Ports for Exchange 2013 taken from https://www.experts-exchange.com/questions/28068143/Exchange-2013-required-ports.html

Requires the AD DS PowerShell module to be installed.

Writes output to both the host and to a text file

.PARAMETER OutputDirectory
Mandatory parameter where the output file will be written to. File is always given the name hostname_yyyyMMdd_HHmmss.txt

.EXAMPLE
C:\Build\Test-ExchangeServerConfiguration.ps1 -OutputDirectory C:\Build\Output

Runs the script, and writes to a file called Hostname_Date_Time.txt in C:\Build\Output
#>

[CmdletBinding(
    SupportsShouldProcess = $True
)]


Param (
    [Parameter(Mandatory=$True)]
    [ValidateScript(
        {
            if(Test-Path $_) {
		        return $true
            } else {
                New-Item $_ -itemtype directory
            }
        }
    )]
    [String]$OutputDirectory

)


#Function to write to both the console and a text file - more reliable than Tee-Object, which sometimes doesn't output as expected
Function Output-Text {
    Param(
        [String]$Filepath,
        [String]$Text,
        [ConsoleColor]$Colour
    )

    if (-not $Colour) {$Colour = "White"}
    
    Write-Host $Text -ForegroundColor $Colour
    $Text |Add-Content -Path $Filepath -Encoding unicode
}

#List of ports to check for connectivity
$ExchangePorts = @("80","443","25","135","26") #Exchange Ports are not used currently - included for a future iteration where they get tidied up - currently hard-coded further down
#$DCPorts = @("389","3268","88","135","636","3269")
$DCPorts = @("389","3268","88","135")
$DNSPorts = @("53")

#List of existing Exchange servers, for Exchange connectivity tests
$ExchangeServers = @("TestEX1",
"TestEx2",
"TestEx3"
)

#List of new servers to be built, for simple PING tests, with FQDN
$HybridServers = @("TestEX2",
"TestEx2",
"TestEx3"
)

#List of DAG Replication Network IP addresses of all new nodes, for PING test over DAG network
$DAGNetIPs = @("10.0.0.6",
"10.0.0.7",
"10.0.0.8"
)

#List of external addresses to resolve and test for connectivity
#As per Office 365 URLs (see header for URL)
$EOLEndPoints = @(
    [pscustomobject]@{host="outlook.office365.com";port=80},
    [pscustomobject]@{host="outlook.office365.com";port=443},
    [pscustomobject]@{host="asl.configure.office.com";port=80},
    [pscustomobject]@{host="asl.configure.office.com";port=443},
    [pscustomobject]@{host="tds.configure.office.com";port=443},
    [pscustomobject]@{host="mshrcstorageprod.blob.core.windows.net";port=80},
    [pscustomobject]@{host="mshrcstorageprod.blob.core.windows.net";port=443},
    [pscustomobject]@{host="tenant.mail.protection.outlook.com";port=25}
)

#SSL Cert thumbprint - replace with Thumbprint for your Exchange Certificate
$Thumbprint = "01234156789abcdef0123456789abcdef01234567"

#File to which output will be written (e.g. server_20171104.txt)
$LogFile = "$OutputDirectory\$(& hostname)_$(get-date -Format yyyyMMdd_HHmmss).txt"
New-Item -ItemType File -Path $LogFile -Force

#*****************Retrieve Net Routing Info****************
Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "Server Routing Table" -Colour Cyan
route print |Add-Content $LogFile
route print

#***************Hybrid Servers CAS interface***************

Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "Hybrid Servers - Client Access Interface PING" -Colour Cyan
Output-Text -Filepath $LogFile -Text $HybridServers

#Run PING test on CAS interface (resolved by DNS)
$HybridPingResults = ForEach ($HybridServer in $HybridServers) {
    Test-NetConnection -ComputerName $HybridServer -ea SilentlyContinue -WarningAction SilentlyContinue -Debug:$DebugPreference
}

#Log servers that PINGed to file
Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "Servers that responded to Ping" -Colour Green

$HybridPingResults |Where-Object {$_.pingsucceeded -eq $True} | Format-Table ComputerName,RemoteAddress,PingSucceeded -AutoSize |Tee-Object -FilePath $LogFile -Append

#Log servers that didn't PING to file, or log that no servers were unresponsive
If (($HybridPingResults |Where-Object {$_.pingsucceeded -eq $False}).count -eq 0) {
    Output-Text -Filepath $LogFile -Text "All Servers Pinged Successfully" -Colour Yellow 
} else {
    Output-Text -Filepath $LogFile -Text "Ping Failed" -Colour Red
    $HybridPingResults |Where-Object {$_.pingsucceeded -eq $False} | Format-Table ComputerName,RemoteAddress,PingSucceeded -AutoSize |Tee-Object -FilePath $LogFile -Append
}

#***************Hybrid Servers DAG interface***************

Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "Hybrid Servers - DAG Interface PING" -Colour Cyan
Output-Text -Filepath $LogFile -Text $DAGNetIPs

#Run PING test to DAG IP Range
$DAGNetIPPingResults = ForEach ($DAGNetIP in $DAGNetIPs) {
    Test-NetConnection -ComputerName $DAGNetIP -ea SilentlyContinue -WarningAction SilentlyContinue -Debug:$DebugPreference
}

#Successful Pinged Servers
Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "Servers that responded to Ping" -Colour Green

$DAGNetIPPingResults |Where-Object {$_.pingsucceeded -eq $True} | Format-Table ComputerName,RemoteAddress,PingSucceeded -AutoSize |Tee-Object -FilePath $LogFile -Append

#Unsuccessful Pinged Servers
If (($DAGNetIPPingResults |Where-Object {$_.pingsucceeded -eq $False}).count -eq 0) {
    Output-Text -Filepath $LogFile -Text "All Servers Pinged Successfully" -Colour Yellow 
} else {
    Output-Text -Filepath $LogFile -Text "Ping Failed" -Colour Red
    $DAGNetIPPingResults |Where-Object {$_.pingsucceeded -eq $False} | Format-Table ComputerName,RemoteAddress,PingSucceeded -AutoSize |Tee-Object -FilePath $LogFile -Append
}

#************************EXCHANGE**************************
#Connectivity and TCP tests to existing Exchange Servers
Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "Existing Exchange Servers" -Colour Cyan 

Output-Text -Filepath $LogFile -Text $ExchangeServers

Output-Text -Filepath $LogFile -Text "Pinging Exchange Servers" -Colour Cyan
$ExchangePingResults = foreach ($ExchangeServer in $ExchangeServers) {
    Test-NetConnection -ComputerName $ExchangeServer -ea SilentlyContinue -WarningAction SilentlyContinue -Debug:$DebugPreference
} 

#Successful Pinged Servers
Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "Servers that responded to Ping" -Colour Green

$ExchangePingResults |Where-Object {$_.pingsucceeded -eq $True} | Format-Table ComputerName,RemoteAddress,PingSucceeded -AutoSize |Tee-Object -FilePath $LogFile -Append

#Unsuccessful Pinged Servers
If (($ExchangePingResults |Where-Object {$_.pingsucceeded -eq $False}).count -eq 0) {
    Output-Text -Filepath $LogFile -Text "All Servers Pinged Successfully" -Colour Yellow 
} else {
    Output-Text -Filepath $LogFile -Text "Ping Failed" -Colour Red
    $ExchangePingResults |Where-Object {$_.pingsucceeded -eq $False} | Format-Table ComputerName,RemoteAddress,PingSucceeded -AutoSize |Tee-Object -FilePath $LogFile -Append
}

#Exchange Port Tests, 80,443,25,26,135, for servers that pinged successfully
Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "Testing Ports..." -Colour Cyan
$ExchangeResults = @()

Output-Text -Filepath $LogFile -Text "HTTP" -Colour Cyan
$ExchangeResults += foreach ($ExchangeServer in ($ExchangePingResults |Where-Object {$_.pingsucceeded -eq $True})) {
    Test-NetConnection -ComputerName $ExchangeServer.ComputerName -Port 80 -ea SilentlyContinue -WarningAction SilentlyContinue -Debug:$DebugPreference
} 

Output-Text -Filepath $LogFile -Text "HTTPS" -Colour Cyan
$ExchangeResults += foreach ($ExchangeServer in ($ExchangePingResults |Where-Object {$_.pingsucceeded -eq $True})) {
    Test-NetConnection -ComputerName $ExchangeServer.ComputerName -Port 443 -ea SilentlyContinue -WarningAction SilentlyContinue -Debug:$DebugPreference
} 

Output-Text -Filepath $LogFile -Text "SMTP" -Colour Cyan
$ExchangeResults += foreach ($ExchangeServer in ($ExchangePingResults |Where-Object {$_.pingsucceeded -eq $True})) {
    Test-NetConnection -ComputerName $ExchangeServer.ComputerName -Port 25 -ea SilentlyContinue -WarningAction SilentlyContinue -Debug:$DebugPreference
}

Output-Text -Filepath $LogFile -Text "Secure SMTP" -Colour Cyan 
$ExchangeResults += foreach ($ExchangeServer in ($ExchangePingResults |Where-Object {$_.pingsucceeded -eq $True})) {
    Test-NetConnection -ComputerName $ExchangeServer.ComputerName -Port 26 -ea SilentlyContinue -WarningAction SilentlyContinue -Debug:$DebugPreference
} 

Output-Text -Filepath $LogFile -Text "RPC EndPoint" -Colour Cyan
$ExchangeResults += foreach ($ExchangeServer in ($ExchangePingResults |Where-Object {$_.pingsucceeded -eq $True})) {
    Test-NetConnection -ComputerName $ExchangeServer.ComputerName -Port 135 -ea SilentlyContinue -WarningAction SilentlyContinue -Debug:$DebugPreference
}

#Output passed and failed connectivity tests
Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "Exchange Connectivity Tests" -Colour Cyan
Output-Text -Filepath $LogFile -Text "Passed" -Colour Green
$ExchangeResults |Where-Object {$_.TcpTestSucceeded -eq $true} |Sort Computername,RemotePort |FT TCPTestSucceeded,ComputerName,RemotePort,RemoteAddress -AutoSize |Tee-Object -FilePath $LogFile -Append

Output-Text -Filepath $LogFile -Text "Failed" -Colour Red
If (($ExchangeResults |Where-Object {$_.TcpTestSucceeded -eq $false}).count -gt 0) {
    $ExchangeResults |Where-Object {$_.TcpTestSucceeded -eq $false} |Sort Computername,RemotePort |FT TCPTestSucceeded,ComputerName,RemotePort,RemoteAddress -AutoSize |Tee-Object -FilePath $LogFile -Append
} else {
    Output-Text -Filepath $LogFile -Text "None" -Colour Green
}

#**************************DCs****************************
Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "Domain Controllers" -Colour Cyan

#Get AD Site of current server from AD
$Site = (Get-ADDomainController -Discover).site
Output-Text -Filepath $LogFile -Text "This server is in AD Site $Site" -Colour Cyan

#Get all Domain Controllers from AD
$DomainControllers = Get-ADDomainController -Filter *

Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "Listing all Domain Controllers in site $Site" -Colour Cyan

#Filter DCs for current Site
$SiteDomainControllers = $DomainControllers |Where-Object {$_.Site -eq $Site}
foreach ($SiteDomainController in $SiteDomainControllers) {
    Output-Text -Filepath $LogFile -Text $SiteDomainController.HostName
}

#Retrieve Sites linked to current Site by SiteLink in AD
$LinkedSites = Get-ADObject -LDAPFilter '(objectClass=siteLink)' -SearchBase (Get-ADRootDSE).ConfigurationNamingContext -Property Name, Cost, Description, Sitelist | `
    Where {$_.sitelist -match $Site} | `
        Select-Object -ExpandProperty SiteList |`
            Sort-Object -Unique | `
                foreach {
                    ($_.split(",")[0]).split("=")[1]
                }
Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "Linked Sites from AD" -Colour Cyan
Output-Text -Filepath $LogFile -Text $LinkedSites 

#Retrieve one DC from each linked Site
$LinkedSiteDomainControllers = foreach ($LinkedSite in $LinkedSites) {
    Get-ADDomainController -SiteName $LinkedSite -Discover
}
Output-Text -Filepath $LogFile -Text "Picking one DC in each Linked Site" -Colour Cyan
Output-Text -Filepath $LogFile -Text $LinkedSiteDCs

#Ping tests for DCs in current Site
Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "Pinging In-Site DCs" -Colour Cyan
$SiteDCPingResults = foreach ($SiteDomainController in $SiteDomainControllers) {
    Test-NetConnection -ComputerName $SiteDomainController.hostname -ea SilentlyContinue -WarningAction SilentlyContinue -Debug:$DebugPreference
} 

#Successful Pinged Servers
Output-Text -Filepath $LogFile -Text "Servers that responded to Ping" -Colour Green

$SiteDCPingResults |Where-Object {$_.pingsucceeded -eq $True} | Format-Table ComputerName,RemoteAddress,PingSucceeded -AutoSize |Tee-Object -FilePath $LogFile -Append

#Unsuccessful Pinged Servers
If (($SiteDCPingResults |Where-Object {$_.pingsucceeded -eq $False}).count -eq 0) {
    Output-Text -Filepath $LogFile -Text "All In-Site DCs Pinged Successfully" -Colour Yellow 
} else {
    Output-Text -Filepath $LogFile -Text "Ping Failed to the following In-Site DCs:" -Colour Red
    $SiteDCPingResults |Where-Object {$_.pingsucceeded -eq $False} | Format-Table ComputerName,RemoteAddress,PingSucceeded -AutoSize |Tee-Object -FilePath $LogFile -Append
}

#TCP Connectivity tests for each DC in current Site
$SiteDCResults = foreach ($Port in $DCPorts) {
    foreach ($SiteDC in $SiteDomainControllers) {
        Test-NetConnection -ComputerName $SiteDC.HostName -Port $Port -ea SilentlyContinue -WarningAction SilentlyContinue -Debug:$DebugPreference
    }
}

#Log results to file
Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "In-Site DC Connectivity Tests" -Colour Cyan
Output-Text -Filepath $LogFile -Text "Passed" -Colour Green
$SiteDCResults |Where-Object {$_.TcpTestSucceeded -eq $True} |Sort Computername,RemotePort |FT TCPTestSucceeded,ComputerName,RemotePort,RemoteAddress -AutoSize |Tee-Object -FilePath $LogFile -Append

Output-Text -Filepath $LogFile -Text "Failed" -Colour Red
If (($SiteDCResults |Where-Object {$_.TcpTestSucceeded -eq $False}).count -gt 0) {
    $SiteDCResults |Where-Object {$_.TcpTestSucceeded -eq $False} |Sort Computername,RemotePort |FT TCPTestSucceeded,ComputerName,RemotePort,RemoteAddress -AutoSize |Tee-Object -FilePath $LogFile -Append
} else {
    Output-Text -Filepath $LogFile -Text "None" -Colour Green
}

#Repeat tests for DCs in linked Sites
Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "Pinging Linked Site DCs" -Colour Cyan
$LinkedSiteDCPingResults = foreach ($LinkedSiteDomainController in $LinkedSiteDomainControllers) {
    Test-NetConnection -ComputerName $LinkedSiteDomainController.hostname -ea SilentlyContinue -WarningAction SilentlyContinue -Debug:$DebugPreference
} 
#Successful Pinged Servers
Output-Text -Filepath $LogFile -Text "Servers that responded to Ping" -Colour Green

$LinkedSiteDCPingResults |Where-Object {$_.pingsucceeded -eq $True} | Format-Table ComputerName,RemoteAddress,PingSucceeded -AutoSize |Tee-Object -FilePath $LogFile -Append

#Unsuccessful Pinged Servers
If (($LinkedSiteDCPingResults |Where-Object {$_.pingsucceeded -eq $False}).count -eq 0) {
    Output-Text -Filepath $LogFile -Text "All Linked Site DCs Pinged Successfully" -Colour Yellow 
} else {
    Output-Text -Filepath $LogFile -Text "Ping Failed to the following Linked Site DCs:" -Colour Red
    $LinkedSiteDCPingResults |Where-Object {$_.pingsucceeded -eq $False} | Format-Table ComputerName,RemoteAddress,PingSucceeded -AutoSize |Tee-Object -FilePath $LogFile -Append
}

$LinkedSiteDCResults = foreach ($Port in $DCPorts) {
    foreach ($LinkedSiteDC in $LinkedSiteDomainControllers) {
        Test-NetConnection -ComputerName $LinkedSiteDC.HostName -Port $Port -ea SilentlyContinue -WarningAction SilentlyContinue -Debug:$DebugPreference
    }
}

Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "Linked Site DC Connectivity Tests" -Colour Cyan
Output-Text -Filepath $LogFile -Text "Passed" -Colour Green
$LinkedSiteDCResults |Where-Object {$_.TcpTestSucceeded -eq $True} |Sort Computername,RemotePort |FT TCPTestSucceeded,ComputerName,RemotePort,RemoteAddress -AutoSize |Tee-Object -FilePath $LogFile -Append

Output-Text -Filepath $LogFile -Text "Failed" -Colour Red
If (($LinkedSiteDCResults |Where-Object {$_.TcpTestSucceeded -eq $False}).count -gt 0) {
    $LinkedSiteDCResults |Where-Object {$_.TcpTestSucceeded -eq $False} |Sort Computername,RemotePort |FT TCPTestSucceeded,ComputerName,RemotePort,RemoteAddress -AutoSize |Tee-Object -FilePath $LogFile -Append
} else {
    Output-Text -Filepath $LogFile -Text "None" -Colour Green
}


#****************Time & DCDiag*******************
$TimeStatus = Invoke-Expression "w32tm /query /status"
Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "Status of time configuration:" -Colour Cyan
Output-Text -Filepath $LogFile -Text "$TimeStatus" -Colour Cyan

$SiteDCDiagDC = ($SiteDomainControllers[0]).HostName
$SiteDCDiagStatus = Invoke-Expression "dcdiag /s:$SiteDCDiagDC"
Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "In-site DCDiag:" -Colour Cyan
Output-Text -Filepath $LogFile -Text "$SiteDCDiagStatus" -Colour Cyan



$RootDCDiagDC = ($DomainControllers | Where-Object {$_.site -eq "Default-First-Site-Name"} |Select-Object -First 1).Hostname
$RootDCDiagStatus = Invoke-Expression "dcdiag /s:$RootDCDiagDC"
Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "Root DCDiag:" -Colour Cyan
Output-Text -Filepath $LogFile -Text "$RootDCDiagStatus" -Colour Cyan

#*****************DNS Servers********************

#Retrieve DNS servers from current server configuration
$DNSServers = Get-DnsClientServerAddress -AddressFamily ipv4 |select -expand serveraddresses -Unique

#test connectivity to port 53 - no PING test as PING may be disabled
$DNSResults = foreach ($Port in $DNSPorts) {
    foreach ($DNSServer in $DNSServers) {
        Test-NetConnection -ComputerName $DNSServer -Port $Port -ea SilentlyContinue -WarningAction SilentlyContinue -Debug:$DebugPreference
    }
}

#Log to file
Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "DNS Connectivity Tests" -Colour Cyan
Output-Text -Filepath $LogFile -Text "Passed" -Colour Green
$DNSResults |Where-Object {$_.TcpTestSucceeded -eq $True} |Sort Computername,RemotePort |FT TCPTestSucceeded,ComputerName,RemotePort,RemoteAddress -AutoSize |Tee-Object -FilePath $LogFile -Append

Output-Text -Filepath $LogFile -Text "Failed" -Colour Red
If (($DNSResults |Where-Object {$_.TcpTestSucceeded -eq $False}).count -gt 0) {
    $DNSResults |Where-Object {$_.TcpTestSucceeded -eq $False} |Sort Computername,RemotePort |FT TCPTestSucceeded,ComputerName,RemotePort,RemoteAddress -AutoSize |Tee-Object -FilePath $LogFile -Append
} else {
    Output-Text -Filepath $LogFile -Text "None" -Colour Green
}


#************EOL EndPoint Tests****************
#Test connectivity to endpoints listed in Office 365 URLs guidance
Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "Exchange Online EndPoint Tests" -Colour Cyan

$EOLEndPointResults = foreach ($EOLEndPoint in $EOLEndPoints) {
        Test-NetConnection -ComputerName $EOLEndPoint.Host -Port $EOLEndPoint.Port -ea SilentlyContinue -WarningAction SilentlyContinue -Debug:$DebugPreference
}

#And log results
Output-Text -Filepath $LogFile -Text "Passed" -Colour Green
$EOLEndPointResults |Where-Object {$_.TcpTestSucceeded -eq $True} |Sort Computername,RemotePort |FT TCPTestSucceeded,ComputerName,RemotePort,RemoteAddress -AutoSize |Tee-Object -FilePath $LogFile -Append

Output-Text -Filepath $LogFile -Text "Failed" -Colour Red
If (($EOLEndPointResults |Where-Object {$_.TcpTestSucceeded -eq $False}).count -gt 0) {
    $EOLEndPointResults |Where-Object {$_.TcpTestSucceeded -eq $False} |Sort Computername,RemotePort |FT TCPTestSucceeded,ComputerName,RemotePort,RemoteAddress -AutoSize |Tee-Object -FilePath $LogFile -Append
} else {
    Output-Text -Filepath $LogFile -Text "None" -Colour Green
}


#************DNS Resolution Check****************
#Test DNS Resolution to Outlook.Office365.com to validate DC it resolves to

Output-Text -FilePath $LogFile -Text ""
Output-Text -FilePath $LogFile -Text "Resolving Outlook.Office365.com to Datacentre" -Colour Cyan
$OutlookHosts = Resolve-DnsName "outlook.office365.com" |Where-Object {$_.QueryType -eq "A"} 
$ResolvedOutlookHostName = ($OutlookHosts |Group Name).Name
Output-Text -Filepath $LogFile -Text "Host Outlook.Office365.com resolved to $ResolvedOutlookHostName" -Colour Green
Output-Text -Filepath $LogFile -Text "Testing Connectivity to resolved endpoints"

#Then test connectivity to each endpoint
$OutlookHostResults = ForEach ($OutlookHost in $OutlookHosts) {
    Test-NetConnection -ComputerName $OutlookHost.IPAddress
}
$AverageRoundTrip = ($OutlookHostResults |Select-Object -ExpandProperty PingReplyDetails |?{$_.RoundTripTime -gt 0}|Measure-Object -Property RoundTripTime -Average).Average
#Take round trip time and average for all successful connections
$AverageRoundTrip = "{0:N0}" -f $AverageRoundTrip


Output-Text -Filepath $LogFile -Text "Average Round Trip Time is $($AverageRoundTrip)ms"


#**********Certificates******************
#Validate Certificate is installed correctly

Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "Validating Certificate is in Machine Personal Store"
$Cert = gci Cert:\LocalMachine\my |Where-Object {$_.Thumbprint -eq $Thumbprint}

If ($Cert -is [System.Security.Cryptography.X509Certificates.X509Certificate2]) {
    If ($Cert.HasPrivateKey -eq $true) { #If Cert doesn't contain the Private Key, it is of no use to Exchange
        Output-Text -Filepath $LogFile -Text "Certificate has Private Key" -Colour Green
        Output-Text -Filepath $LogFile -Text "Hostnames in Certificate"
        Output-Text -Filepath $LogFile -Text $cert.DnsNameList.Unicode
    } else {
        Output-Text -Filepath $LogFile -Text "Problem with Certificate Private Key - please verify" -Colour Red
    }
} else {
    Output-Text -Filepath $LogFile -Text "Problem retrieving certificate from Store - please verify" -Colour Red
}

Write-Host -ForegroundColor DarkBlue "Tests Completed, please review the output at $LogFile" -BackgroundColor Gray
