<#
 
.Notes
References to:
https://blogs.technet.microsoft.com/david231/2015/03/30/for-exchange-2010-and-2013-do-this-before-calling-microsoft/
https://blogs.technet.microsoft.com/heyscriptingguy/2012/11/27/use-powershell-and-wmi-or-cim-to-view-and-to-set-power-plans/
#>
 
$PrereqPath = "C:\Build\Pre-reqs"
 
Write-Host "Setting power plan" -ForegroundColor Cyan
$PowerPlan = Get-CimInstance -Namespace root\cimv2\power -Class win32_PowerPlan | Where-Object {$_.ElementName -eq "High Performance"}
If ($PowerPlan.IsActive -ne $True) {
    Invoke-CimMethod -InputObject $PowerPlan -MethodName Activate
    Write-Host "High Performance power plan set as active" -ForegroundColor Green
} else {
    Write-Host "High Performance power plan already set as active" -ForegroundColor Cyan
}
 
 
If (-not (Test-Path "HKLM:\Software\Policies\Microsoft\Windows NT\Rpc")) {
    New-Item "HKLM:\Software\Policies\Microsoft\Windows NT" -Name "Rpc"
}
 
#Check minimum TCP Connection timeout and set inline with recommendation if not already
if ((Get-ItemProperty "HKLM:\Software\Policies\Microsoft\Windows NT\RPC\") -match "MinimumConnectionTimeout") { #Test if MinimumConnectionTimeout exists
    if ((Get-ItemProperty "HKLM:\Software\Policies\Microsoft\Windows NT\RPC\").MinimumConnectionTimeout -eq '0x00000078') { #And check the value if it does
        Write-Host "MinimumConnectionTimeout matches recommended value (120 decimal)" -ForegroundColor Green
    } else { #Else for property value - anything other than 120 Decimal enters here
        Write-Host "Setting Property" -ForegroundColor Yellow
        Set-ItemProperty "HKLM:\Software\Policies\Microsoft\Windows NT\RPC\" -Name MinimumConnectionTimeout -Value '0x00000078'
    }
} else { #Else for property not existing - so we'll create it
    if (New-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\Windows NT\RPC\" -Name "MinimumConnectionTimeout" -PropertyType DWord -Value '0x00000078') {
        Write-Host "MinimumConnectionTimeout property created successfully" -ForegroundColor Green
    } else {
        Write-Error "Error creating MinimumConnectionTimeout property"
    }
}
 
#Check TCP Keepalive and set it to appropriate value (20 minutes)
if ((Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters") -match "KeepAliveTime") { #Test if KeepAliveTime exists
    if ((Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters").KeepAliveTime -eq '1200000') { #And check the value if it does
        Write-Host "KeepAliveTime matches recommended value (20 minutes)" -ForegroundColor Green
    } else { #Else for property value - anything other than 20 minutes enters here
        Write-Host "Setting Property" -ForegroundColor Yellow
        Set-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters" -Name KeepAliveTime -Value '1200000'
    }
} else { #Else for property not existing - so we'll create it
    if (New-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters" -Name "KeepAliveTime" -PropertyType DWord -Value '1200000') {
        Write-Host "KeepAliveTime property created successfully" -ForegroundColor Green
    } else {
        Write-Error "Error creating KeepAliveTime property"
    }
}
 
#Set TCP configuration - RSS and Chimney
Invoke-Expression  "netsh int tcp set global rss=enabled" |Out-Null
Write-Host "RSS Enabled" -ForegroundColor Green
Invoke-Expression "netsh int tcp set global chimney=Automatic" |Out-Null
Write-Host "TCP Chimney set to automatic" -ForegroundColor Green
 
 
#This section taken from Install-Exchange15, by Michel de Rooij, michel@eightwone.com
#available from http://eightwone.com
#All servers will have at least 32Gb RAM, so Pagefile needs to be 32GB+10MB
Write-Host "Configuring page file" -ForegroundColor Cyan
$CS = Get-WmiObject -Class Win32_ComputerSystem -EnableAllPrivileges
 
If ($CS.AutomaticManagedPagefile) {
    Write-Verbose 'System configured to use Automatic Managed Pagefile, reconfiguring'
    Try {
        $CS.AutomaticManagedPagefile = $false
        # RAM + 10 MB, with maximum of 32GB + 10MB
        $InstalledMem= $CS.TotalPhysicalMemory
        $DesiredSize= (32GB+10MB)/ 1MB
        $tmp= $CS.Put()
        $CPF= Get-WmiObject -Class Win32_PageFileSetting |Where-Object {$_.Name -eq "C:\Pagefile.sys"}
        $CPF.Delete()
        Set-WmiInstance -Class Win32_PageFileSetting -Arguments @{name="P:\pagefile.sys"; InitialSize = $DesiredSize; MaximumSize = $DesiredSize} -EnableAllPrivileges |Out-Null
        $PF = Get-WmiObject Win32_PageFileSetting -Filter "SettingID='pagefile.sys @ P:'"
        $PF.Name = "P:\Pagefile.sys"
        $PF.InitialSize= $DesiredSize
        $PF.MaximumSize= $DesiredSize
        $PF.Caption = "P:\pagefile.sys"
        #$tmp= $PF.Put()
 
        $CPF= Get-WmiObject -Class Win32_PageFileSetting
        Write-Output "Pagefile set to manual, at $($CPF.Name), initial/maximum size: $($CPF.InitialSize)MB / $($CPF.MaximumSize)MB" 
 
    }
    Catch {
        Write-Error "Problem reconfiguring pagefile: $($ERROR[0])"
    }
}
Else {
    Write-Host 'Manually configured page file, skipping configuration' -ForegroundColor Yellow
}
 
Install-WindowsFeature -Restart AS-HTTP-Activation, `
    Desktop-Experience, `
    NET-Framework-45-Features, `
    RPC-over-HTTP-proxy, `
    RSAT-Clustering, `
    RSAT-Clustering-CmdInterface, `
    RSAT-Clustering-Mgmt, `
    RSAT-Clustering-PowerShell, `
    Web-Mgmt-Console, `
    WAS-Process-Model, `
    Web-Asp-Net45, `
    Web-Basic-Auth, `
    Web-Client-Auth, `
    Web-Digest-Auth, `
    Web-Dir-Browsing, `
    Web-Dyn-Compression, `
    Web-Http-Errors, `
    Web-Http-Logging, `
    Web-Http-Redirect, `
    Web-Http-Tracing, `
    Web-ISAPI-Ext, `
    Web-ISAPI-Filter, `
    Web-Lgcy-Mgmt-Console, `
    Web-Metabase, `
    Web-Mgmt-Console, `
    Web-Mgmt-Service, `
    Web-Net-Ext45, `
    Web-Request-Monitor, `
    Web-Server, `
    Web-Stat-Compression, `
    Web-Static-Content, `
    Web-Windows-Auth, `
    Web-WMI, `
    Windows-Identity-Foundation, `
    RSAT-ADDS, `
    Failover-Clustering, `
    Server-Media-Foundation
 
 
