<#
.SYNOPSIS
Script to move distribution list management from On-prem Exchange to O365

.DESCRIPTION
Takes one or more DG names, disables them locally, creates a contact on-prem and a new DG of the same name in O365

.PARAMETER DGs
One or more Distribution Group names in a comma separated list

.PARAMETER LocalExchangeServer
The FQDN of the On-Premise Exchange server

.PARAMETER LocalDirSyncServer
The FQDN of the On-Premise DirSync server - can be excluded if using the NoDirSync parameter, otherwise will prompt for it

.PARAMETER Tenant
The short name of the tenant in Office365, for example Contoso if the FQDN were contoso.onmicrosoft.com

.PARAMETER ContactsOU
The LDAP DN of the OU to store the new contacts, eg 'OU=MailContacts,DC=Contoso,DC=com'

.PARAMETER DLRoot
Specifies the root folder that will be the location of all the exported DG XML files

.PARAMETER ExcludedFQDNs
Enter any FQDNs that on-premises objects have as email domains that need to be excluded from the Tenant e.g. "contoso.local"

.PARAMETER NoDirSync
Specify this switch to disable automatic syncing using a DirSync server

.EXAMPLE

Can be used to migrate a single Distribution Group:

    Move-DistributionGroupOnline -DGs "DG1" -LocalExchangeServer "EX1.contoso.com" -LocalDirSyncServer "AADS1.contoso.com" -Tenant "Contoso" -ContactsOU "OU=MailContacts,DC=Contoso,DC=com"

.EXAMPLE

Can be used to move multiple Distribution Groups:

    Move-DistributionGroupOnline -DGs "DG1","DG2" -LocalExchangeServer "EX1.contoso.com" -LocalDirSyncServer "AADS1.contoso.com" -Tenant "Contoso" -ContactsOU "OU=MailContacts,DC=Contoso,DC=com"

.EXAMPLE

Can be used inline with a text file containing one DG name per line:

    Move-DistributionGroupOnline -DGs (Get-Content D:\Tools\DGs.txt) -LocalExchangeServer "EX1.contoso.com" -LocalDirSyncServer "AADS1.contoso.com" -Tenant "Contoso" -ContactsOU "OU=MailContacts,DC=Contoso,DC=com"

.EXAMPLE

Can be used to migrate in bulk with a foreach loop:

    Get-DistributionGroup |foreach {Move-DistributionGroupOnline -DGs $_.name -LocalExchangeServer "EX1.contoso.com" -LocalDirSyncServer "AADS1.contoso.com" -Tenant "Contoso" -ContactsOU "OU=MailContacts,DC=Contoso,DC=com"}

Note that this has potentially huge impact!

.EXAMPLE

Can be used with the NoDirSync switch to allow the user to manually initiate the sync with the tenant, if using a custom sync mechanism. If using this switch, the -LocalDirSyncServer Parameter can be omitted.

    Move-DistributionGroupOnline -DGs "DG1" -LocalExchangeServer "EX1.contoso.com" -NoDirSync -Tenant "Contoso" -ContactsOU "OU=MailContacts,DC=Contoso,DC=com"
#>


Param (
    [Parameter(Mandatory=$true,position=0)]
        [string[]]$DGs,
    [Parameter(Mandatory=$true,position=1)]
        [string]$LocalExchangeServer,
    [Parameter(Mandatory=$false,position=2)]
        [string]$LocalDirSyncServer,
    [Parameter(Mandatory=$true,position=3)]
        [string]$Tenant,
    [Parameter(Mandatory=$true,position=4)]
        [string]$ContactsOU,
    [Parameter(Mandatory=$false)]
        [string[]]$ExcludedFQDNs,
    [Parameter(Mandatory=$false)]
        [string]$DLRoot,
    [Switch]$NoDirSync = $False
)

function Select-FolderDialog($Title='Select A Folder', $Directory = 0) {
    $object = New-Object -comObject Shell.Application

    $folder = $object.BrowseForFolder(0, $Title, 0, $Directory)
    if ($folder -ne $null) {
        $folder.self.Path
    }
}

$Host.UI.RawUI.BackgroundColor = "Black"
$Host.UI.RawUI.ForegroundColor = "White"

Clear-Host

$DebugPreference = "SilentlyContinue"
$ErrorAction = "Continue"



if (-not $InitialConfigComplete) { #Using this means that the script can be run repeatedly without having to enter credentials every time


    Write-Debug "Checking Tenant Session exists"
    #Create and import Exchange Online session
    $CheckTenant = Get-Command Get-CloudMailUser -ErrorAction "silentlycontinue"

    if (-not $CheckTenant) {
        $TenantCred = Get-Credential -Message "Please enter the Tenant credentials"
        write-host "Connecting to tenant"
        $TenantSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $TenantCred -Authentication Basic -AllowRedirection
        Import-PSSession $TenantSession -Prefix "Cloud"
        $CheckTenant = Get-Command Get-CloudMailUser
    }

    Write-Debug "Checking local Exchange session exists"
    #Create and import local Exchange session
    $CheckLocalExchange = Get-Command Get-LocalMailUser -ErrorAction "SilentlyContinue"
    if (-not $CheckLocalExchange) {
        $LocalExchangeCred = Get-Credential -Message "Please enter the Exchange Admin credentials"
        $LocalExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$LocalExchangeServer/powershell" -cred $LocalExchangeCred
        Import-PSSession $LocalExchangeSession -Prefix "Local"
        $CheckLocalExchange = Get-Command Get-LocalMailUser
    }

    if (-not $NoDirSync) {
        if (-not $LocalDirSyncServer) {
            $LocalDirSyncServer = Read-Host "Please enter the FQDN of the local DirSync server"
        }

        Write-Debug "Checking Dirsync session exists"
        #Create and import DirSync session
        $CheckDirSync = Get-Command Start-DSOnlineCoexistenceSync -ErrorAction "SilentlyContinue"

        if (-not $CheckDirSync) {
            $DirSyncCred = Get-Credential -Message "Please enter the DirSync credentials"
            Write-Host "Connecting to DirSync Server"
            $DirSyncSession = New-PSSession -ComputerName $LocalDirSyncServer -cred $DirSyncCred
            #Ensure that DirSync comands are available
            Invoke-Command -scriptblock {& "$($env:ProgramFiles)\Windows Azure Active Directory Sync\DirSync\ImportModules.ps1"} -Session $DirSyncSession
            Import-PSSession $DirSyncSession -Prefix "DS" -Module "Microsoft.Online.Coexistence.PS.Config"
            $CheckDirSync = Get-Command Start-DSOnlineCoexistenceSync
        }
    } else {
        $CheckDirSync = $True
    }

    
    
    

    if ($CheckTenant -and $CheckLocalExchange -and $CheckDirSync) {
        Write-Debug "Initial Config OK"
        $InitialConfigComplete = $true
    }
}


if (-not $InitialConfigComplete) {
    Write-Error "Initial Config Failed"
} else
{

    if (-not $DLRoot) { #DLRoot will contain the folder that will hold the exported XML files

        $DLRoot = Select-FolderDialog -Title "Please choose the location to store the DL XML Files"

        if (-not $DLRoot) {
            #Write-Error "Cancel clicked"
            return "Cancel clicked"
        }
    }

        $Time = (get-date).ToString("yyyyMMdd_HHmmss")

        $DLPath = "$DLRoot\ExportedDGs_$Time"
        #Using a timestamp means that the script can be run repeatedly and previous output will be kept

        If (-not (Test-Path $DLPath)) {
            New-Item $DLPath -ItemType Directory
        }


    Write-Debug "Getting local DLs"
    $LocalDistributionGroups = $DGs

    #Clear-Host

    #Prompt for confirmation to continue - important because confirmation for disabling the DLs is set to $false
    Write-Host -BackgroundColor Black -ForegroundColor Red "WARNING: $($LocalDistributionGroups.count) Distribution Groups will be disabled and recreated in Exchange Online - Do you wish to continue? Type YES in upper case to continue"
    $response = Read-Host

    if ($response -cne "YES") { #Means that confirmation wasn't given - a failsafe
        return "Confirmation not given"
    }

    $FailedDLs = @()
    $ConflictingRecipients = @()

    ForEach ($DG in $LocalDistributionGroups) {
    
        $DLExported = $False
        $DLExportPath = "$DLPath\$($DG)"
        if (Test-Path $DLExportPath) {
        } else {
            New-Item $DLExportPath -ItemType Directory |out-null
        }

        Write-Debug "Retrieving on-prem DL $($DG) and members"

        #retrieve DG and members for use later
        $DL = Get-LocalDistributionGroup -Identity $DG
        $DLMembers = Get-LocalDistributionGroupMember -Identity $DG
        
        $DLDN = "LDAP://$($DL.DistinguishedName)"
        $DLADObj = [ADSI]$DLDN
        
        $DLMemberof = @()
        foreach ($MemberOf in $ADObj.MemberOf) {
            $DLMemberof += (Get-LocalDistributionGroup $memberof)
        }

        #Export DG and it's members to XML for logging - this enables the DG to be recreated if needed
        Export-Clixml -InputObject $DL -Path "$DLExportPath\Group.xml"
        Export-Clixml -InputObject (Get-LocalDistributionGroupMember -Identity $DL.name) -Path "$DLExportPath\Members.xml"
        Export-CLIXML -InputObject $DLMemberOf -Path "$DLExportPath\MemberOf.xml"

        $DLExported = ((Test-Path "$DLExportPath\Group.xml") -and (Test-Path "$DLExportPath\Members.xml") -and (Test-Path "$DLExportPath\MemberOf.xml"))

        if ($DLExported) { #means we can continue with this DG
            $CloudDLExists = $true
            Write-Host "Distribution Group $($DL.Name) Exported" -ForegroundColor Green
            write-host "Disabling DG $($DL.Name)" -ForegroundColor Yellow

            Disable-LocalDistributionGroup -Identity $DL.Name -Confirm:$false
            Write-Host "Distribution Group $($DL.Name) Disabled" -fore Red
            
            If ($NoDirSync -eq $False) {
                $DirSyncCount = 0
                do {
                    if ($DirSyncCount -ge 10) {
                        Write-Error "DirSync hasn't completed or DG not removed from the tenant after 5 minutes, please investigate"
                        $Response = Read-Host "Please type YES to continue running the script or NO to quit"
                        if ($response -ceq "YES") {
                        } elseif ($response -ceq "NO") {
                            return
                        }
                    }
                    #initiate sync and wait 30 seconds
                    Start-DSOnlineCoexistenceSync 
                    write-host "Initiating Online Sync and waiting 30 seconds" -fore "Yellow"
                    $DirSyncCount++
                    Start-Sleep -Seconds 30
                
                
                    if (Get-CloudDistributionGroup -Identity $DL.Name -ErrorAction "Silentlycontinue") {#check if DG still exists in O365 and suppress error message
                        Write-Host "Distribution Group still exists in Tenant" -ForegroundColor Yellow
                        $CloudDLExists = $True
                    } else {
                        Write-Host "Distribution Group not found in Tenant" -ForegroundColor Green
                        $CloudDLExists = $False
                    }
                } until ($CloudDLExists -eq $False) #keep repeating until the DL isn't found

            } else {
                
                do {                
                    if (Get-CloudDistributionGroup -Identity $DL.Name -ErrorAction "Silentlycontinue") {#check if DG still exists in O365 and suppress error message
                        $CloudDLExists = $True
                        Write-Host "Distribution Group still exists in Tenant" -ForegroundColor Yellow
                        Write-Host "Please initiate the sync with O365 and verify that the DL is removed from the tenant"
                        do {
                            $Response = Read-Host "Type YES continue or NO to quit the script"
                        } until (($response -ceq "YES") -or ($response -ceq "NO"))
                        if ($response -eq "NO") {
                            Write-Host "User initiated termination of script" -ForegroundColor Red
                            $response = Read-Host 'Are you sure? 
This will result in the local DG being disabled, but not recreated in the tenant
type "YES" to continue'
                            if ($response -ceq "YES") {
                                return
                            }
                        }
                            
                    } else {
                        Write-Host "Distribution Group not found in Tenant" -ForegroundColor Green
                        $CloudDLExists = $False
                    }
                } until ($CloudDLExists -eq $False) #keep repeating until the DL isn't found
                

            }
            

            Write-Debug "Cloud DL removed"
            
            $CloudRecipient = Get-CloudRecipient -Identity "$($dl.alias)@$Tenant.mail.onmicrosoft.com" -ErrorAction "SilentlyContinue" #check if Recipient exists and suppress error message

            if ($CloudRecipient) {

                $ConflictingRecipients += $dl.alias

            } else {

                
                #Create On-Prem contact and set details as per previous DG
                $NewMC = New-LocalMailContact -DisplayName $DL.DisplayName `
                    -Name $DL.Name `
                    -Alias $DL.Alias `
                    -ExternalEmailAddress "$($dl.alias)@$Tenant.mail.onmicrosoft.com" `
                    -PrimarySmtpAddress $dl.PrimarySmtpAddress `
                    -OrganizationalUnit $ContactsOU

                $MCEmailAddresses = $dl.EmailAddresses
                $MCEmailAddresses += "x500:$($dl.LegacyExchangeDN)"

                Set-LocalMailContact -Identity $NewMC.Alias -EmailAddresses $MCEmailAddresses

                ForEach ($DGMemberOf in $DLMemberof) {
                    if (Get-CloudDistributionGroup -Identity $DGMemberOf.Name) {
                        Add-CloudDistributionGroupMember -Member $NewMC.Name -Identity $DGMemberOf.Name
                    }
                    else {
                        Add-LocalDistributionGroupMember -Member $NewMC.Name -Identity $DGMemberOf.Name
                    }
                }

                Write-Host "On-prem Mail contact created" -ForegroundColor Green
                
                $MgtUser = Get-CloudUser ($dl.ManagedBy).split("/")[-1]

                #Create DG in O365 and set values as per previous DG
                $NewDG = New-CloudDistributionGroup -PrimarySmtpAddress $dl.primarysmtpaddress `
                    -Name $DL.Name `
                    -DisplayName $DL.DisplayName `
                    -ModerationEnabled $dl.ModerationEnabled `
                    -ManagedBy $MgtUser.DistinguishedName `
                    -Alias $DL.Alias `
                    -MemberDepartRestriction $DL.MemberDepartRestriction

                Write-Host "Cloud DL created" -ForegroundColor Green

                #Allow time for creation to complete
                start-sleep -Seconds 10

                $CloudEmailAddresses = $dl.EmailAddresses

                foreach ($FQDN in $ExcludedFQDNs) {
                    $CloudEmailAddresses = $CloudEmailAddresses |?{$_ -notlike "*$FQDN"} #filter out SMTP addresses for on-prem FQDN if it's not a valid external domain - remove the where clause if this is not the case
                }

                $CloudEmailAddresses += "x500:$($dl.LegacyExchangeDN)"

                #Add necessary email addresses and also timestamp the DL to match up with exported on-prem DG XML files later on if necessary
                Set-CloudDistributionGroup -identity $newdg.Alias -EmailAddresses $CloudEmailAddresses -CustomAttribute15 $Time

                write-host "Email addresses added" -ForegroundColor Green

                foreach ($DLMember in $DLMembers) { #add DG members using primary SMTP as it should be the same
                    Add-CloudDistributionGroupMember -Identity $NewDG.alias -Member $DLMember.PrimarySMTPAddress
                }

                Write-Host "Members added" -ForegroundColor Green

            }

        } else { #We get to here if the DG didn't export for any reason
            Write-Host "Distribution Group $($DG.Name) failed to export" -ForegroundColor Red
            $FailedDLs += $($DG.name) #add the name to the failed array
            continue
        }

        Write-Host "Distribution Group $($DL.Name) Moved Successfully to Office 365" -BackgroundColor DarkGreen -ForegroundColor White
    }

    Write-Host "DLs that failed to export are listed in $DLPath\FailedDLs.txt"
    Set-Content -Path "$DLPath\FailedDLs.txt" -Value $FailedDLs
    Write-Host "`n"
    Write-Host "Recipients that conflicted are listed in $DLPath\ConflictingRecipients.txt"
    Set-Content -Path "$DLPath\ConflictingRecipients.txt" -Value $ConflictingRecipients

}
