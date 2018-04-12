<#
.DESCRIPTION
Retrieves all Skype for Business Online Client Policies and lists the settings configured by each

.EXAMPLE
    .\Get-SfBOPolicySettings.ps1

#>

#Requires -Modules SkypeOnlineConnector

Param(
    $Credential
)

Function Connect-LyncOnline {
    Param (
        [Parameter(Mandatory=$True,position=0)]
            [ValidateScript({$_ -is [System.Management.Automation.PSCredential]})]
            [System.Management.Automation.PSCredential]$Credential
    )

    $LyncSession = New-CsOnlineSession -Credential $Credential
    Import-PSSession $LyncSession -Prefix LO
}

If (-not ((Get-Command -ListImported) -match "Get-LocsOnlineUser")) {

    Connect-LyncOnline -Credential $Credential

}

$ClientPolicies = Get-LOCsClientPolicy |Select-Object Identity,disable*,enable*

$Props = $ClientPolicies |Get-Member disable*,enable* |Select-Object name

foreach ($Policy in $ClientPolicies) {
    
    Write-Host $Policy.Identity -ForegroundColor Cyan
    foreach ($Prop in $Props) {
        $Name = $Prop.name
        #$name
        
        if ((-not $($Policy.$Name)) -and ($($Policy.$Name).count -gt 0)) {
            
            Write-Host "$Name : " -NoNewline
            Write-Host $($Policy.$name) -ForegroundColor Yellow
        }
        if ($Policy.$Name) {
            
            Write-Host "$Name : " -NoNewline
            Write-Host $($Policy.$name) -ForegroundColor Green
        }
    }
    Write-Host "`n"

} 

