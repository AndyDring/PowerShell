<#
.SYNOPSIS

Script 

.DESCRIPTION
Takes a text file containing domain names and adds them to a tenancy, generating a text file containing the DNS verification records in the same directory. 
Can also complete the verification process.
Written by Andy Dring, October 2016
andy@andydring.it

.PARAMETER Domains
List of domains

.PARAMETER UseContoso
Switch to show sample output for contoso.com

#>

Param(
    [Parameter(position=0)]
        [Array]$Domains
)

Clear-Host 

If (-not $Domains) {
    $Domains = "contoso.com"
}


foreach ($Domain in $Domains) {

    Write-Host "Checking " -NoNewline
    Write-Host "$Domain`n" -ForegroundColor Yellow

    $SipRecord = "sip.$Domain"
    $Sip = Resolve-DnsName -Type ANY -Name $sipRecord -Server 8.8.8.8 -ErrorAction SilentlyContinue

    $LyncDiscoverRecord = "lyncdiscover.$Domain"
    $LyncDiscover = Resolve-DnsName -Type ANY -Name $LyncDiscoverRecord -Server 8.8.8.8 -ErrorAction SilentlyContinue

    $SipFedRecord = "_sipfederationtls._tcp.$Domain"
    $SipFed = Resolve-DnsName -Type ANY -Name $SipFedRecord -Server 8.8.8.8 -ErrorAction SilentlyContinue

    $SipTLSRecord = "_sip._tls.$Domain"
    $SipTLS = Resolve-DnsName -Type ANY -Name $SipTLSRecord -Server 8.8.8.8 -ErrorAction SilentlyContinue


    if ($Sip) {
        write-host "$SipRecord points to:" -ForegroundColor Green
        ForEach ($Record in $Sip) {
            If ($Record.Type -eq "CNAME") {
                Write-Host "CNAME: $($Record.NameHost)" -ForegroundColor Green
            } elseif (($Record.Type -eq "A") -or ($Record.Type -eq "AAAA")) {
                Write-Host "$($Record.Type): $($Record.IPAddress)" -ForegroundColor Green
            } else {
                Write-Host "Unexpected record type - please check manually using:" -ForegroundColor Yellow
                Write-Host "Resolve-DnsName -Type ANY -Name $SipRecord -Server 8.8.8.8" -ForegroundColor Yellow
            }
        }
    } else {
        Write-Host "$SipRecord not found" -ForegroundColor Red
    }

    if ($LyncDiscover) {
        write-host "`n$LyncDiscoverRecord points to:" -ForegroundColor Green 
        ForEach ($Record in $LyncDiscover) {
            If ($Record.Type -eq "CNAME") {
                Write-Host "CNAME: $($Record.NameHost)" -ForegroundColor Green
            } elseif (($Record.Type -eq "A") -or ($Record.Type -eq "AAAA")) {
                Write-Host "$($Record.Type): $($Record.IPAddress)" -ForegroundColor Green
            } else {
                Write-Host "Unexpected record type - please check manually using" -ForegroundColor Yellow
                Write-Host "Resolve-DnsName -Type ANY -Name $LyncDiscoverRecord -Server 8.8.8.8" -ForegroundColor Yellow
            }
        }
    } else {
        Write-Host "`n$LyncDiscoverRecord not found" -ForegroundColor Red
    }

    if ($SipFed) {
        write-host "`n$SipFedRecord points to:" -ForegroundColor Green
        foreach ($Record in $SipFed) {
            If ($Record.Type -eq "SRV") {
                Write-Host "$($Record.NameTarget) on Port $($Record.Port) with Priority $($Record.Priority), Weight $($Record.Weight)" -ForegroundColor Green
            } else {
                Write-Host "Unexpected record type - please check manually using" -ForegroundColor Yellow
                Write-Host "Resolve-DnsName -Type ANY -Name $SipFedRecord -Server 8.8.8.8" -ForegroundColor Yellow 
            }
        }
    } else {
        Write-Host "`n$SipFedRecord not found" -ForegroundColor Red
    }

    if ($SipTLS) {
        Write-Host "`n$SipTLSRecord points to:" -ForegroundColor Green
        Foreach ($Record in $SipTLS) {
            If ($Record.Type -eq "SRV") {
                Write-Host "$($Record.NameTarget) on Port $($Record.Port) with Priority $($Record.Priority), Weight $($Record.Weight)" -ForegroundColor Green
            } else {
                Write-Host "Unexpected record type - please check manually using" -ForegroundColor Yellow
                Write-Host "Resolve-DnsName -Type ANY -Name $SipTLSRecord -Server 8.8.8.8" -ForegroundColor Yellow 
            }        }
    } else {
        Write-Host "`n$SipTLSRecord not found" -ForegroundColor Red
    }

    Clear-Variable sip,lyncdiscover,sipfed,siptls
    Write-Host "`n"
}

