<#
.SYNOPSIS

Script to add domains to O365/Azure tenancy and generate verification records

.DESCRIPTION
Takes a text file containing domain names and adds them to a tenancy, generating a text file containing the DNS verification records in the same directory. 
Can also complete the verification process.
Written by Andy Dring, November 2014
andy@andydring.it

.PARAMETER TenantCredential
Takes a credential with appropriate Admin rights in the tenancy

.PARAMETER Complete
When specified, the script will complete the verification, instead of adding the domains. If DomainList is specified, it will verify the domains in the list, otherwise it will attempt to verify all unverified domains in the tenant

.PARAMETER DomainFile
Text file containing list of domains to add

.PARAMETER GenerateRecords
Suppresses addition of new domains, only generates records for unverified domains

#>

Param(
    [Parameter(Mandatory=$True,position=0)]
        [ValidateScript({$_ -is [System.Management.Automation.PSCredential]})]
        [System.Management.Automation.PSCredential]$TenantCredential,
    [Switch]$Complete = $False,
    [Parameter(Mandatory=$False)]
        [ValidateScript({Test-Path $_})]
        [String]$DomainFile,
    [Switch]$GenerateOnly = $False
)

<#
If (-not $TenantCredential) {

        $TenantCredential = Get-Credential -Message "Please enter the Tenant adminstrative credentials"

}
#>

Connect-MsolService -Credential $TenantCredential

Function Select-FileDialog
{
	param([string]$Title=$args[0],[string]$Directory=0,[string]$Filter="All Files (*.*)|*.*")
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	$objForm = New-Object System.Windows.Forms.OpenFileDialog
	$objForm.InitialDirectory = $Directory
	$objForm.Filter = $Filter
	$objForm.Title = $Title
	$objForm.ShowHelp = $true
	$Show = $objForm.ShowDialog()
	If ($Show -eq "OK")
	{
		Return $objForm.FileName
	}
	Else
	{
		Write-Error "Operation cancelled by user."
        break
	}
}

function Select-FolderDialog($Title='Select A Folder', $Directory = 0) {  
    $object = New-Object -comObject Shell.Application   
      
    $folder = $object.BrowseForFolder(0, $Title, 0, $Directory)  
    if ($folder -ne $null) {  
        $folder.self.Path  
    }  
} 


If ($Complete -eq $False) {

    If ($GenerateOnly -eq $False) { #In other words, we want to add the domains as well as generate the records

        If (-not $DomainFile) { 
            $DomainFile = Select-FileDialog -Title "Please select the file containing the list of domains" -Filter "Text Files (*.txt)|*.txt"
        }
        
        $NewDomains = Get-Content $DomainFile

        ForEach ($NewDomain in $NewDomains) {
	        New-MSOLDomain -Name $NewDomain
        }

        Start-Sleep -Seconds 5

    }

    If ($DomainFile) {

        $UnverifiedDomains = Get-Content $DomainFile 

        $VerificationRecords = @()

        ForEach ($UnverifiedDomain in $UnverifiedDomains) {

            $Record = Get-MsolDomainVerificationDns -DomainName $UnverifiedDomain
    
            $Obj = New-Object System.Object
    
            $Value = $Record.Label.split(".")[0]
            $Obj |Add-Member -MemberType NoteProperty -Name "RecordType" -Value "TXT"
            $Obj |Add-Member -MemberType NoteProperty -Name "Alias" -Value "@ or $($UnverifiedDomain)"
            $Obj |Add-Member -MemberType NoteProperty -Name "Value" -Value "MS=$Value"
            $Obj |Add-Member -MemberType NoteProperty -Name "TTL" -Value "3600"

            $VerificationRecords += $Obj

        }

    } else {

        $UnverifiedDomains = Get-MsolDomain |Where-Object {$_.Status -eq "Unverified"} 

        $VerificationRecords = @()

        ForEach ($UnverifiedDomain in $UnverifiedDomains) {

            $Record = Get-MsolDomainVerificationDns -DomainName $UnverifiedDomain.Name
    
            $Obj = New-Object System.Object
    
            $Value = $Record.Label.split(".")[0]
            $Obj |Add-Member -MemberType NoteProperty -Name "RecordType" -Value "TXT"
            $Obj |Add-Member -MemberType NoteProperty -Name "Alias" -Value "@ or $($UnverifiedDomain.Name)"
            $Obj |Add-Member -MemberType NoteProperty -Name "Value" -Value "MS=$Value"
            $Obj |Add-Member -MemberType NoteProperty -Name "TTL" -Value "3600"

            $VerificationRecords += $Obj

        }

    }

    If (-not $DomainFile) {
        $Folder = Select-FolderDialog -Title "Please choose the directory to store the output file"
        $OutputPath = "$($Folder)\DNSVerificationRecords.txt"
    } else {
        $OutputPath = "$(Split-Path $DomainFile -parent)\DNSVerificationRecords.txt"
    }

        $VerificationRecords | Sort Alias |Format-Table -AutoSize | Tee-Object -FilePath $OutputPath
        Write-Host -ForegroundColor Green "Records written to $OutputPath"
 
} else {
    #Completes verification for the domains in the file

    If ($DomainFile) {
        foreach ($Domain in (Get-Content $DomainFile)) {
            Confirm-MsolDomain -DomainName $Domain
        }
    } else {

        ForEach($Domain in (Get-MsolDomain |Where-Object {$_.Status -eq "Unverified"} )) {
            Confirm-MsolDomain -DomainName $Domain.Name
        }

    }
}

