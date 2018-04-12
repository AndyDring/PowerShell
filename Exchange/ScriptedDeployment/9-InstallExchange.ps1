$ServerName = $(hostname)

$ExInstallFolder = "c:\build\ex2013cu19"


$InstallCommand = "$ExInstallFolder\Setup.exe /Mode:Install /Roles:Mailbox,ClientAccess,ManagementTools /OrganizationName:ADT /TargetDir:`"E:\Exchange`" /LogFolderPath:`"E:\MDB\$Servername-InstallDB\Log`" /mdbname:$Servername-InstallDB /dbfilepath:`"E:\MDB\$Servername-InstallDB\$hostname-InstallDB.edb`" /IAcceptExchangeServerLicenseTerms"
Invoke-Expression $InstallCommand

#notepad c:\ExchangeSetupLogs\ExchangeSetup.log