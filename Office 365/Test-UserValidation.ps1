#Requires -module MSOnline

<#
.NOTES
Written by Andy Dring, March 2018
andy@andydring.it


.SYNOPSIS
Script to check validity of users for migration to Exchange Online. Checks against on-premises Exchange, Azure AD and Exchange Online

.DESCRIPTION
Takes a list of email addresses in a CSV file, with the header 'mail', finds the user in Exchange Online, Exchange On-Premises and Azure AD and assesses the suitability of the found objects for migration to Exchange Online.

Script exits if a successful connection to all of on-premises Exchange, Exchange Online and Azure AD cannot be made.

Checks (in order):
Object exists in Exchange On-Premises
On-premises Object is a Mailbox
Object with the same UPN exists in Azure AD
Azure AD User is Licenced for Exchange
On-premises Primary SMTP matches the UPN
Object is valid in Exchange Online
On-premises Exchange GUID matches Exchange Online GUID
Primary SMTP address doesn't contain any characters that aren't allowed in Exchange Online
On-premises object contains an Exchange Online routing address (@tenant.mail.onmicrosoft.com)
Checks that the routing address prefix matches the primary SMTP address prefix

Once all users are checked, sends two emails to the recipients hard-coded into the script, one with the invalid users CSV as an attachment, the other with the valid users CSV.

.PARAMETER OutputDirectory
Specifies the directory that will be used for output files


.PARAMETER UserCSV
Specifies the path to the CSV file containing the list of email addresses to be assessed


.PARAMETER SMTPServer
FQDN  of the SMTP server to use to send mail messages containing output


.PARAMETER SMTPCredentials
Optionally provide Credentails to use against the SMTP server sending the email. Can be left blank, in which case the connection will be done under the context running the script. 
If specified, must contain a PSCredential object


.OUTPUTS
Outputs the DistinguishedName, DisplayName and the Mail attribute to a CSV file called UserValidation-ValidUsers-<Time/Date stamp>.csv

Outputs the email address provided and a reason for failure to a CSV file called User-Validation-InvalidUsers-<Time/Date stamp>.csv

The <Time/Date> stamp is in the format yyyyMMdd-HHmmss


.EXAMPLE
C:\Migration\Test-UserValidation.ps1 -OutputDirectory C:\Migration\Output -UserCSV C:\Migration\Users.csv -SMTPServer Exchange1.contoso.com

Outputs two files to the specified directory, reading in from the specified CSV file. Sends emails via Exchange1.contoso.com, as the user account running the script.


.EXAMPLE
C:\Migration\Test-UserValidation.ps1 -OutputDirectory C:\Migration\Output -UserCSV C:\Migration\Users.csv -SMTPServer Exchange1.contoso.com -SMTPCredentials (Get-Credential Contoso\ExchAdmin)

Outputs two files to the specified directory, reading in from the specified CSV file. Prompts for a password and sends emails via Exchange1.contoso.com, with the supplied credentials.
#>

Param(
    [Parameter(Mandatory=$True)]
        [ValidateScript(
            {
                Test-Path $_ 
            }
        )]
        [String]$OutputDirectory,
    [Parameter(Mandatory=$True)]
        [ValidateScript(
            {
                Test-Path $_ 
            }
        )]
        [String]$UserCSV,
    [Parameter(Mandatory=$True)]
        [String]$SMTPServer,
    [Parameter(Mandatory=$False)]
        [PSCredential]$SMTPCredentials
)

#Function to output text to both a log file and the screen simultaneously
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

#Variable declaration
New-Variable -Name DateStamp -Value (get-date -Format yyyyMMdd-HHmmss) -Option Constant
New-Variable -Name LogFile -Value ("$OutputDirectory\UserValidation-$DateStamp.log") -Option Constant
New-Variable -Name ValidUserFile -Value ("$OutputDirectory\UserValidation-$DateStamp-ValidUsers.csv")
New-Variable -Name InvalidUserFile -Value ("$OutputDirectory\UserValidation-$DateStamp-InvalidUsers.csv")
New-Variable -Name TenantName -Value "contoso" -Option Constant
New-Variable -Name Recipients -Value @("servicedesk@contoso.com","sdm@contoso.com") -Option Constant
New-Variable -Name FromAddress -Value "O365.Migrations@contoso.com" -Option Constant

#List of characters to check within prefix of SMTP addresses
#Be happy to have a better way of handlig this
$InvalidCharacterSet = @('~',
'!',
'@',
'#',
'$',
'%',
'£',
'^',
'&',
'*',
'(',
')',
'-',
'+',
'=',
'[',
']',
'{',
'}',
'\',
'/',
'|',
';',
':',
'<',
'>',
'?',
',',
'"',
"'",
'¦'
"``")

#Initialisation Section
Clear-Host

#Import MSOnline cmdlets
If ([string](Get-Module) -notmatch "MSOnline") {
    Import-Module MSOnline
}

#Import on-premises Exchange session
If (Get-Command Get-OPUser -ErrorAction SilentlyContinue) {
    Write-Host "On-premises Exchange session available" -ForegroundColor Green
} Else { #On-premises Exchange cmdlets not available, so importing session
    Write-Host "On-premises Exchange session not available, connecting..." -ForegroundColor Cyan
    $ExchangeServer = Read-Host "Please enter the FQDN of the Exchange Server to connect to"
    $Cred = Get-Credential
    $ExSession= New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ -Authentication Kerberos -Credential $Cred -ErrorAction SilentlyContinue
    If ($ExSession) {
        Try {
            Write-Host "Importing on-premises Exchange session with cmdlet prefix OP" -ForegroundColor Green
            Import-PSSession $ExSession -Prefix OP -ErrorAction SilentlyContinue
        } Catch {
            
            Write-Warning "Unable to connect to on-premises Exchange. Please verify details and try again"
            $_
            break
        }
    } Else {
        Write-Warning "Unable to connect to on-premises Exchange. Please verify details and try again"
        break
    }
}

#Import Exchange Online session - note that this requires an App Password for the account being used

If (Get-Command Get-XORecipient -ErrorAction SilentlyContinue) {
    Write-Host "Exchange Online session available" -ForegroundColor Green
} Else { #Exchange Online cmdlets not available, so importing session
    Write-Host "Exchange Online session not available, connecting..." -ForegroundColor Cyan
    $NewExOModule = (Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse ).FullName

    If ($NewExOModule) {
        Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse |Sort-Object LastWriteTime -Descending).FullName|?{$_ -notmatch "_none_"}|Select-Object -First 1)
        $ExOSession = New-ExoPSSession -ErrorAction SilentlyContinue
    } else {
        Write-Warning "Exchange Online MFA module is unavailable. This can be installed by following the instructions at https://technet.microsoft.com/en-us/library/mt775114(v=exchg.160).aspx" 
        Write-Host "Would you like to launch this now? " -ForegroundColor Yellow -NoNewline
        If ((Read-Host) -eq "Yes") {
            Start-Process "https://technet.microsoft.com/en-us/library/mt775114(v=exchg.160).aspx"
            break
        } Else {
            If ((Read-Host "If your account has MFA configured, to connect Exchange Online requires an App Password. Type `'Yes`' to proceed") -eq "Yes") {
        
                $o365Cred= Get-Credential -Message "Please enter the tenant credentials"
                $ExOSession= New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $o365Cred -AllowRedirection -Authentication Basic -ErrorAction SilentlyContinue
            } Else {
                Write-Warning "Unable to connect to Exchange Online. Please verify details and try again"
                break
            }
        }
    }
    
    If ($ExOSession) {
        Try {
            Write-Host "Importing Exchange Online session with cmdlet prefix XO" -ForegroundColor Green
            Import-PSSession $ExOSession -Prefix XO -ErrorAction SilentlyContinue 
        } Catch {
            
            Write-Warning "Unable to connect to Exchange Online. Please verify details and try again"
            $_
            break
        }
    } Else {
        Write-Warning "Unable to connect to Exchange Online. Please verify details and try again"
        break
    }


}

#Connect to Azure AD
If (Get-MsolUser -ErrorAction SilentlyContinue) {
    Write-Host "Azure AD cmdlets available" -ForegroundColor Green
} Else {
    Write-Host "Azure AD cmdlets not available, connecting..." -ForegroundColor Cyan
    $Error.clear()
    Connect-MsolService -ErrorAction SilentlyContinue
    If (-not $Error) {
        Write-Host "MSOnline cmdlets imported" -ForegroundColor Green  
    } Else {
        Write-Warning "Unable to import MSOnline cmdlets. Please verify and try again"
        break
    }
}

Output-Text -Filepath $LogFile -Text "All modules initialised and ready" -Colour Green

Output-Text -Filepath $LogFile -Text "Importing users from file $UserCSV" -Colour Cyan

$Users = Import-Csv -Path $UserCSV 

Output-Text -Filepath $LogFile -Text "`n$($Users.count) users imported from file"

$ValidUsers = @()
$InvalidUsers = @("`"EmailAddress`",`"Reason`"")
$InvalidDomains = @("BadDomain1.com","BadDomain2.com")

Try {
    $MSOLUsers = Get-MsolUser -All -ErrorAction SilentlyContinue
} Catch {
    "Unable to retrieve users from Azure AD - exiting script " |Add-Content -Path $LogFile -Encoding unicode
    Write-Error -Message "Unable to retrieve Azure AD users"
    break
} 

Try {
    $MSOLDomains = Get-MsolDomain -ErrorAction SilentlyContinue
} Catch {
    "Unable to retrieve domains from Azure AD - exiting script" |Add-Content -Path $LogFile -Encoding unicode
    Write-Error -Message "Unable to retrieve Azure AD domains"
    break
}


foreach ($User in $Users) {

    Output-Text -Filepath $LogFile -Text "`nProcessing user $($User.mail)" -Colour White

   
    #Find any mail-enabled object
    $OPUser = Get-OPUser -Identity $User.mail -ErrorAction SilentlyContinue

    #If 
    If (-not $OPUser) {
        Output-Text -FilePath $LogFile -Text "Invalid user - Email not found" -Colour Red
        $InvalidUsers += @("`"$($User.mail)`",`"EmailNotFound`"")
        Continue
    } Else {
        Output-Text -Filepath $LogFile -Text "User found at `n$($OPUser.DistinguishedName)`n" -Colour Cyan
    }

    #Determine if it's a mailbox
    If ($OPUser.RecipientType -ne "UserMailbox") {
        Output-Text -FilePath $LogFile -Text "Invalid user - object is not a mailbox" -Colour Red
        $InvalidUsers += @("`"$($User.mail)`",`"NotMailbox`"")
        Continue
    } Else {
        Write-Verbose "User $($OPUser.Name) is a mailbox"
        $OPMailbox = Get-OPMailbox -Identity $OPUser.DistinguishedName
    }

    #Check if the user exists in AzureAD - if not everything else is irrelevant
    $MSOLUser = $MSOLUsers |Where-Object {$_.UserPrincipalName -eq $OPUser.UserPrincipalName} -ErrorAction SilentlyContinue

    If (-not $MSOLUser) {
        Output-Text -FilePath $LogFile -Text "Invalid user - user not found in AzureAD" -Colour Red
        $InvalidUsers += @("`"$($User.mail)`",`"MSOLUserNotFound`"")
        Continue
    } else {
        Write-Verbose "User $($MSOLUser.DisplayName) found in AzureAD"
        $Licenced = $MSOLUser.Licenses.servicestatus |Where {($_.serviceplan.servicetype -eq "Exchange") -and ( $_.ProvisioningStatus -eq "Success")}

        If (-not $Licenced) {
            Output-Text -FilePath $LogFile -Text "Invalid user - user not licenced for Exchange in AzureAD" -Colour Red
            $InvalidUsers += @("`"$($User.mail)`",`"UserNotLicenced`"")
            Continue
        } Else {
            Write-Verbose "User has the following Exchange licences `n"
            Write-Verbose $Licenced
        }
    }

    $OPPrimarySMTP = ($OPMailbox.EmailAddresses |Where-Object {$_ -cmatch "SMTP:"}).split(":")[1]
    Write-Verbose "Primary SMTP is $OPPrimarySMTP"

    If ($OPMailbox.UserPrincipalName -eq $OPPrimarySMTP) {
        Write-Verbose "User UPN and Primary SMTP match"
        $XORecipient = Get-XORecipient -Identity $OPPrimarySMTP -ErrorAction SilentlyContinue
        If (-not $XORecipient) {
            Output-Text -FilePath $LogFile -Text "Invalid user - AzureAD UPN mismatch to Primary SMTP" -Colour Red
            $InvalidUsers += @("`"$($User.mail)`",`"AzureADUPNMismatch`"")
            Continue  
        } else {
            Write-Verbose "User $($XOMailbox.DistinguishedName) found in Exchange Online"
            If ($OPMailbox.ExchangeGUID -ne $XORecipient.ExchangeGUID) {
                Output-Text -FilePath $LogFile -Text "Invalid user - ExchangeGUID mismatch" -Colour Red
                $InvalidUsers += @("`"$($User.mail)`",`"ExchangeGUIDMismatch`"")
                Continue                  
            }
        }
    } Else {
        Output-Text -FilePath $LogFile -Text "Invalid user - on-prem UPN mismatch to Primary SMTP" -Colour Red
        $InvalidUsers += @("`"$($User.mail)`",`"OnPremUPNMismatch`"")
        Continue
    }
 
    Foreach ($EmailAddress in ($OPMailbox.EmailAddresses |Where-Object {($_ -match "smtp") -and ($_ -notmatch ".local")})) {
        $SMTPDomain = $EmailAddress.split("@")[1]
        Write-Verbose "SMTP Domain is $SMTPDomain"
        If ($MSOLDomains.name -match $SMTPDomain) {
            Write-Verbose "SMTP Domain $SMTPDomain matched Verified Domain"
        } elseif ($InvalidDomains -match $SMTPDomain) {
            Output-Text -FilePath $LogFile -Text "Invalid user - has email address of `@$($InvalidDomains -match $SMTPDomain) that doesn't match a domain in the tenant" -Colour Red
            $InvalidUsers += @("`"$($User.mail)`",`"NonMigrateableEmailDomain`"")
            Continue
        } else {
            Output-Text -FilePath $LogFile -Text "Invalid user - has routable email address that doesn't match a domain in the tenant" -Colour Red
            $InvalidUsers += @("`"$($User.mail)`",`"NonTenantDomain`"")
            Continue
        }

        $UniqueEmailAddressPart = ($EmailAddress.split("@")[0]).Split(":")[1]
        Write-Verbose "Unique email address part is:$UniqueEmailAddressPart"
        foreach ($Character in $UniqueEmailAddressPart) {
            If ($InvalidCharacterSet -contains $Character) {
                Output-Text -FilePath $LogFile -Text "Invalid user - Invalid character in SMTP prefix" -Colour Red
                $InvalidUsers += @("`"$($User.mail)`",`"InvalidEmailPrefix`"")
                Continue
            }
        }

        If ($UniqueEmailAddressPart -like ".*") {
            Output-Text -FilePath $LogFile -Text "Invalid user - Invalid character in SMTP prefix" -Colour Red
            $InvalidUsers += @("`"$($User.mail)`",`"InvalidEmailPrefix`"")
            Continue
        }

        If ($UniqueEmailAddressPart -like "*.") {
            Output-Text -FilePath $LogFile -Text "Invalid user - Invalid character in SMTP prefix" -Colour Red
            $InvalidUsers += @("`"$($User.mail)`",`"InvalidEmailPrefix`"")
            Continue
        }

        If ($UniqueEmailAddressPart -like "..") {
            Output-Text -FilePath $LogFile -Text "Invalid user - `'..`' in SMTP prefix" -Colour Red
            $InvalidUsers += @("`"$($User.mail)`",`"InvalidEmailPrefix`"")
            Continue
        }
    }

    $TenantMailAddress = $OPMailbox.EmailAddresses |Where-Object {$_ -match "$TenantName.mail.onmicrosoft.com"}
    Write-Verbose "Tenant routing address is $TenantMailAddress"

    If (-not $TenantMailAddress) {
        Output-Text -FilePath $LogFile -Text "Invalid user - on-prem user has no tenant routing address" -Colour Red
        $InvalidUsers += @("`"$($User.mail)`",`"NoTenantRoutingAddress`"")
        Continue
    } else {
        $PrimarySMTPPrefix = $OPPrimarySMTP.split("@")[0]
        Write-Verbose "PrimarySMTP prefix is $PrimarySMTPPrefix"
        $TenantMailAddressPrefix = ($TenantMailAddress.split("@")[0]).split(":")[1]
        Write-Verbose "Tenant routing prefix is $TenantMailAddressPrefix"

        If ($PrimarySMTPPrefix -ne $TenantMailAddressPrefix) {
            Output-Text -FilePath $LogFile -Text "Invalid user - routing address prefix mismatch to Primary SMTP" -Colour Red
            $InvalidUsers += @("`"$($User.mail)`",`"RoutingAddressPrefixMismatch`"")
            Continue
        }
    }

    Output-Text -Filepath $LogFile -Text "User $($User.mail) was valid" -Colour Green
    $ValidUsers += $User |select distinguishedname,mail,displayname
}

Set-Content -Path $InvalidUserFile -Value $InvalidUsers
#Import-Csv $OutputDirectory\InvalidUsers-$DateStamp.csv

$ValidUsers | Export-Csv -Path  $ValidUserFile -NoTypeInformation

Try {
    Write-Verbose "Sending invalid users CSV to Recipients list"
    
    If ($SMTPCredentials) {
            Send-MailMessage -Attachments $InvalidUserFile `
            -Body "List of invalid users from file $UserCSV" `
            -BodyAsHtml -From $FromAddress `
            -Subject "Invalid Users" `
            -To $Recipients `
            -SmtpServer $SMTPServer `
            -Credential $SMTPCredentials
    } Else {
        Send-MailMessage -Attachments $InvalidUserFile `
            -Body "List of invalid users from file $UserCSV" `
            -BodyAsHtml -From $FromAddress `
            -Subject "Invalid Users" `
            -To $Recipients `
            -SmtpServer $SMTPServer
    }

} Catch {
    Output-Text -Filepath $LogFile -Text "`nFailed to send Invalid Users email to recipient list - $($_.Exception.Message)"
    Write-Error "Failed to send Invalid Users email to recipient list - $($_.Exception.Message)"
}

Try {
    Write-Verbose "Sending valid users CSV to Recipients list"
    
    If ($SMTPCredentials) {
            Send-MailMessage -Attachments $ValidUserFile `
            -Body "List of valid users from file $UserCSV" `
            -BodyAsHtml -From $FromAddress `
            -Subject "Valid Users" `
            -To $Recipients `
            -SmtpServer $SMTPServer `
            -Credential $SMTPCredentials
    } Else {
        Send-MailMessage -Attachments $ValidUserFile `
            -Body "List of valid users from file $UserCSV" `
            -BodyAsHtml -From $FromAddress `
            -Subject "Valid Users" `
            -To $Recipients `
            -SmtpServer $SMTPServer
    }

} Catch {
    Output-Text -Filepath $LogFile -Text "`nFailed to send Valid Users email to recipient list - $($_.Exception.Message)"
    Write-Error "Failed to send Valid Users email to recipient list - $($_.Exception.Message)"
}