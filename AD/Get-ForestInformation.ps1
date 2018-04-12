Import-Module ActiveDirectory



<#
$ApplicationPartitions
$DomainNamingMaster
$ForestDomains
$ForestMode
$GCs 
$ForestName 
$SchemaMaster
$Sites 
$UPNSuffixes$domaindeta
#>
#$ForestDetails = @()

$Forest = Get-ADForest 

$ApplicationPartitions = $Forest |Select-Object -ExpandProperty ApplicationPartitions
$DomainNamingMaster = $Forest |Select-Object -ExpandProperty DomainNamingMaster
$ForestDomains = $Forest |Select-Object -ExpandProperty Domains
$ForestMode = $Forest |Select-Object -ExpandProperty ForestMode
$GCs = $Forest |Select-Object -ExpandProperty GlobalCatalogs
$ForestName = $Forest |Select-Object -ExpandProperty Name
$SchemaMaster = $Forest |Select-Object -ExpandProperty SchemaMaster
$Sites = $Forest |Select-Object -ExpandProperty Sites
$UPNSuffixes = $Forest |Select-Object -ExpandProperty UPNSuffixes

$ForestDetails = New-Object System.Object
$ForestDetails |Add-Member -Name "Name" -MemberType NoteProperty -Value $ForestName
$ForestDetails |Add-Member -Name "FunctionalLevel" -MemberType NoteProperty -Value $ForestMode
$ForestDetails |Add-Member -Name "DomainNamingMaster" -MemberType NoteProperty -Value $DomainNamingMaster
$ForestDetails |Add-Member -Name "SchemaMaster" -MemberType NoteProperty -Value $SchemaMaster
$ForestDetails |Add-Member -Name "GCs" -MemberType NoteProperty -Value $GCs
$ForestDetails |Add-Member -Name "Sites" -MemberType NoteProperty -Value $Sites
$ForestDetails |Add-Member -Name "UPNSuffixes" -MemberType NoteProperty -Value $UPNSuffixes

$DomainDetails = @()

foreach ($ForestDomain in $ForestDomains) {
    $Domain = Get-ADDomain $ForestDomain -Server $ForestDomain

    If ($Domain) {

        $DomObj = New-Object System.Object

        $DomObj |Add-Member -Name "DomainName" -MemberType NoteProperty -Value $Domain.Name.ToString()
        $DomObj |Add-Member -Name "NetBIOSName" -MemberType NoteProperty -Value $Domain.NetBIOSName.ToString()
        $DomObj |Add-Member -Name "DistinguishedName" -MemberType NoteProperty -Value $Domain.DistinguishedName.ToString()
        $DomObj |Add-Member -Name "FQDN" -MemberType NoteProperty -Value $Domain.DNSRoot.ToString()
        $DomObj |Add-Member -Name "FunctionalLevel" -MemberType NoteProperty -Value $Domain.DomainMode.ToString()
        $DomObj |Add-Member -Name "InfrastructureMaster" -MemberType NoteProperty -Value $Domain.InfrastructureMaster.ToString()
        $DomObj |Add-Member -Name "PDCEmulator" -MemberType NoteProperty -Value $Domain.PDCEmulator.ToString()
        $DomObj |Add-Member -Name "RIDMaster" -MemberType NoteProperty -Value $Domain.RIDMaster.ToString()
        $DCs = @()
        ForEach ($Replica in ($Domain | Select-Object -ExpandProperty ReplicaDirectoryServers)) {
            $DC = Get-ADDomainController $Replica
            $DCOS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Replica
            $DCCS = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Replica

            $DCObj = New-Object System.Object

            $DCObj |Add-Member -Name "FQDN" -MemberType NoteProperty -Value $DC.HostName.ToString()
            $DCObj |Add-Member -Name "DistinguishedName" -MemberType NoteProperty -Value $DC.ComputerObjectDN.ToString()
            $DCObj |Add-Member -Name "IPAddress" -MemberType NoteProperty -Value $DC.IPv4Address.ToString()
            $DCObj |Add-Member -Name "ADSite" -MemberType NoteProperty -Value $DC.Site.ToString()
            $DCObj |Add-Member -Name "GC" -MemberType NoteProperty -Value $DC.IsGlobalCatalog.ToString()
            $DCObj |Add-Member -Name "OSName" -MemberType NoteProperty -Value ($DCOS.Name).Split("|")[0]
            $DCObj |Add-Member -Name "OSBuild" -MemberType NoteProperty -Value $DCOS.Version.ToString()
            if ($DCOS.OSArchitecture) {
                $DCObj |Add-Member -Name "OSArchitecture" -MemberType NoteProperty -Value $DCOS.OSArchitecture.ToString()
            } else {
                $DCObj |Add-Member -Name "OSArchitecture" -MemberType NoteProperty -Value "x86"
            }
        
            $DCObj |Add-Member -Name "Memory" -MemberType NoteProperty -Value $DCCS.TotalPhysicalMemory.ToString()
            $DCObj |Add-Member -Name "Cores" -MemberType NoteProperty -Value (((Get-WMIObject -Class Win32_Processor -ComputerName $Replica) |Select-Object -First 1).NumberOfCores * $DCCS.NumberOfProcessors)
            $DCObj |Add-Member -Name "Manufacturer" -MemberType NoteProperty -Value $DCCS.Manufacturer.ToString()
            $DCObj |Add-Member -Name "Model" -MemberType NoteProperty -Value $DCCS.Model.ToString()

            $DCs += $DCObj
        }
        $DomObj |Add-Member -Name "DCs" -MemberType NoteProperty -Value $DCs

        $DomainDetails += $DomObj

    } else {
        
        $DomObj = New-Object System.Object

        $DomObj |Add-Member -Name "DomainName" -MemberType NoteProperty -Value "$ForestDomain not contactable, with error $($Error[0].Exception)"

        $DomainDetails += $DomObj
    }

    Clear-Variable DomObj
}

$ForestDetails |Add-Member -Name "Domains" -MemberType NoteProperty -Value $DomainDetails

$ForestDetails 

