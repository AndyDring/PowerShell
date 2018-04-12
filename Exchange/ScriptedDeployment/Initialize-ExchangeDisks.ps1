#Requires -version 4.0
#Requires -RunAsAdministrator

<#
.SYNOPSIS
Uses RA disks presented to server to create mount points for Exchange installation volume, Database mount points and Restore Disk.
Includes an option to clear disks if needed - THIS WILL ERASE ALL DATA ON THOSE DISKS
Thoroughly test this script before deploying on servers that matter to you

.DESCRIPTION
Takes input CSV of databases. See the DBFile parameter for more information.

Initialises Exchange install disk to E:, and database and restore disks as mountpoints under folder hard-coded in script.
Exchange disk is detected by a size of 1TB. Database disks are detected by a size of 1150GB. Restore disk is detected by a size of 900GB.
Modifying these value in the script allows for them to be customised to your environment. 
Hardcoding them means they do not need to be specified everytime the script is run, resulting in unnecessary parameters.

All disks are initialised as GPT, and will only be initialised if either the disks are currently RAW or if the ClearDisks switch is specified.

If the number of DB disks does not match the number of databases retrieved from the CSV, the script will exit with an error.

Exchange install disk initialisation can be skipped with the SkipExchangeInstall switch.

Outputs a logfile with the name FQDN_MountpointConfiguration_yyyyMMdd_HHmmss.txt" to the OutputDirectory.

.PARAMETER OutputDirectory
Specifies the directory to which the log file will be written. If it doesn't exist, the script creates it

.PARAMETER DBFile
CSV file containing the databases that this server will host. It is expected that the server will be one of many, using a reasonably standard DAG design, so the CSV file format is:

Name      Primary      Secondary      Third      Lagged
DB01      Exch1        Exch2          Exch7      Exch8
DB02      Exch2        Exch1          Exch6      Exch5
DB03      Exch6        Exch7          Exch5      Exch2
DB04      Exch8        Exch4          Exch2      Exch9

The Name heading is mandatory, the other header names don't matter. 
The script matches the DB name if that row contains the hostname of the server running the script.
In the example above, running on server Exch5 would result in a DB list containing DB02,DB03 and running on server Exch2 would result in a DB list of DB01,DB02,DB03,DB04

.PARAMETER ClearDisks
In order to configure the mountpoints, by default the disks must be RAW. If they are not, this switch instructs the script to wipe existing disks.
IF THIS OPTION IS SPECIFIED, ANY DATA ON THE DISKS SELECTED WILL BE DESTROYED 
It will not wipe anything marked as a System Disk, and ignores Disk 0, which is typically the Windows installation disk.
The script will present a list of disks to be cleared, and will prompt for each disk to be cleared.
This is about as step-by-step as it can be written, but extreme care should be taken when using this switch to prevent data loss

.PARAMETER SkipExchangeInstall
If the disk to which Exchange will be installed is already configured, this switch will prevent the script from attempting to re-initialise it.

.EXAMPLE
C:\Build\Initialize-ExchangeDisks.ps1 -OutputDirectory C:\Build\output -DBFile C:\Build\Databases.csv 
Configures the Exchange disks using C:\Build\Databases.csv to name the volumes and mountpoints, and writing the log file to C:\Build\Output\exch1.domain.com_MountpointConfiguration_20180325_134655.txt

.EXAMPLE
C:\Build\Initialize-ExchangeDisks.ps1 -OutputDirectory C:\Build\output -DBFile C:\Build\Databases.csv -ClearDisks
Erases any MBR disks that are not marked as System, and then configures the Exchange disks using C:\Build\Databases.csv to name the volumes and mountpoints, and writing the log file to C:\Build\Output\exch1.domain.com_MountpointConfiguration_20180325_134655.txt.


.EXAMPLE
C:\Build\Initialize-ExchangeDisks.ps1 -OutputDirectory C:\Build\output -DBFile C:\Build\Databases.csv -SkipExchangeInstall 
Configures the Exchange disks using C:\Build\Databases.csv to name the volumes and mountpoints, but doesn't initialise the Exchange install disk.

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
                New-Item $_ -itemtype Directory
            }
        }
    )]
    [String]$OutputDirectory,
    [Parameter(Mandatory=$False)]
    [ValidateScript(
        {
            if(Test-Path $_) {
		        return $true
            } else {
                Write-Error "Directory $_, specified as MountRoot doesn't exist"
                Break
            }
        }
    )]
    [String]$MountRoot = "E:\MDB",

    [Parameter(Mandatory=$True)]
    [ValidateScript(
        {
            Test-Path $_
        }
    )]
    [String]$DBFile,
    [Switch]$ClearDisks,
    [Switch]$SkipExchangeInstall
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

Function Clear-Disks {

    If ($SkipExchangeInstall) {
        Output-Text -Filepath $logfile -Text ""
        Output-Text -Filepath $logfile -Text "Skipping Exchange Install Disk" -Colour Yellow
        $Disks = Get-Disk |?{($_.IsSystem -ne $true) -and ($_.Number -ne 0) -and ($_.PartitionStyle -ne "RAW") -and ($_.Size -ne $ExchangeDiskSize)} 

    } else {

        $Disks = Get-Disk |?{($_.IsSystem -ne $true) -and ($_.Number -ne 0) -and ($_.PartitionStyle -ne "RAW")} 
    }
    Output-Text -Filepath $logfile -Text ""
    Output-Text -Filepath $logfile -Text "The following disks will be cleared:" -Colour Yellow
    $Disks |Sort-Object Number |Select-Object Number,Size,PartitionStyle,Path |Format-Table -AutoSize

    $Disks |Sort-Object Number |Select-Object Number,Size,PartitionStyle,Path |Format-Table -AutoSize >> $logfile

    $Disks |Clear-Disk -RemoveData

    Read-Host "Disks cleared, press Enter to continue"
} 


$Delay = 1
$LogFile = "$OutputDirectory\$((gwmi win32_ComputerSystem).DNSHostName).$((gwmi win32_ComputerSystem).Domain)_MountpointConfiguration_$(Get-Date -Format yyyyMMdd_HHmmss).txt"
#$MountRoot = "E:\MDB"
Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "Mountpoint Root is $MountRoot" -Colour Cyan

$DBDiskSize = 250GB
$RestoreDiskSize = 150GB
$ExchangeDiskSize = 100GB

If ($ClearDisks) {
    Output-Text -Filepath $LogFile -Text ""
    If ($SkipExchangeInstall) {
        Output-Text -Filepath $LogFile -Text "Warning: Option to clear all non-RAW, non-system disks except Exchange Install was specified. `n`nAll data on disks will be erased! `n`nType Yes to continue" -Colour Yellow
    } Else {
        Output-Text -Filepath $LogFile -Text "Warning: Option to clear all non-RAW, non-system disks was specified. `n`nAll data on disks will be erased! `n`nType Yes to continue" -Colour Yellow
    } 
    If ((Read-Host) -eq "Yes") {
        Output-Text -Filepath $LogFile -Text "`'Yes`' response received, clearing disks" -Colour Yellow
        Clear-Disks
    } else {
        Output-Text -Filepath $LogFile -Text "`'Yes`' response not received. exiting script" -Colour Red
        break
    }
} else {
    Output-Text -Filepath $LogFile -Text ""
    Output-Text -Filepath $LogFile -Text "ClearDisks switch not specified, skipping clearing disks" -Colour Cyan
}

#**********EXCHANGE INSTALLATION***********
if ($SkipExchangeInstall) {
    Output-Text -Filepath $LogFile -Text ""
    Output-Text -Filepath $LogFile -Text "Option to skip Exchange Install Disk initialisation selected, skipping" -Colour Yellow
} else {

    Output-Text -Filepath $LogFile -Text ""
    Output-Text -Filepath $LogFile -Text "Initialising Exchange Volume" -Colour Cyan

    $ExchangeInstallDisk = Get-Disk | Where-Object {($_.PartitionStyle -eq "Raw") -and ($_.Size -eq $ExchangeDiskSize)}
    If (($ExchangeInstallDisk |measure).count -gt 1) {
        Output-Text -Filepath $LogFile -Text "More than one disk of size $ExchangeDiskSize found, choosing first disk for Exchange Install Disk" -Colour Yellow
        $ExchangeInstallDisk = $ExchangeInstallDisk[0]
    } elseif (($ExchangeInstallDisk |measure).count -lt 1) {
        Output-Text -Filepath $LogFile -Text "Less than one free disk of size $ExchangeDiskSize found, skipping Exchange Install Disk" -Colour Yellow
    } else {

        Output-Text -Filepath $LogFile -Text "Creating Exchange disk using Disk $($ExchangeInstallDisk.Number)" -Colour Cyan
        Initialize-Disk -Number $ExchangeInstallDisk.Number -PartitionStyle GPT
        $Partition = New-Partition -DiskNumber $ExchangeInstallDisk.Number -UseMaximumSize 
        Start-Sleep -Seconds $Delay
        $Partition |Format-Volume -FileSystem NTFS -NewFileSystemLabel "Exchange" -Confirm:$false
        $Partition |Add-PartitionAccessPath -AccessPath "E:"
        # -DiskNumber $ExchangeInstallDisk.Number -PartitionNumber $Partition.PartitionNumber 

        If ((Get-Partition).AccessPaths -contains "E:\") {
            Output-Text -Filepath $LogFile -Text "Exchange volume initialised at E:\" -Colour Green
        } else {
            Output-Text -Filepath $LogFile -Text "Failure initialising Exchange volume at E:\" -Colour Red
        }

    }
    If (Test-Path $MountRoot) {
        Output-Text -Filepath $LogFile -Text "Mountpoint root $MountRoot already exists" -Colour Yellow
    } else {
        if (Test-Path (Split-Path $MountRoot -Parent)) {
            Output-Text -Filepath $LogFile -Text "Creating mountpoint root $MountRoot" -Colour Green
            New-Item $MountRoot -ItemType Directory
        } else {
            Output-Text -Filepath $LogFile -Text "Parent path for Mountpoint root ($(Split-Path $MountRoot -Parent)) doesn't exist - unable to continue" -Colour Yellow
            break 
        }
    }
}

Clear-Variable Partition -ErrorAction SilentlyContinue
#**************DATABASE DISKS**************
$CSV = $DBFile

$Databases = ipcsv $csv 

$DBList = ($Databases |?{$_ -match $((gwmi win32_ComputerSystem).DNSHostName)}).name
Output-Text -Filepath $LogFile -Text "Databases for host $(hostname) are $DBList" -Colour Cyan
$DBResponse = Read-Host "If this list is correct, type Yes to continue or anything else to abort"
If ($DBResponse -ne "Yes") {
    Output-Text -Filepath $LogFile -Text "Yes response not received, aborting" -Colour Cyan
    break
} else {
    Output-Text -Filepath $LogFile -Text "Yes response received, proceeding" -Colour Cyan
}


Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "Initialising Database Disks" -Colour Cyan

If (Test-Path $MountRoot) {
    Output-Text -Filepath $LogFile -Text "Mountpoint root $MountRoot already exists" -Colour Yellow
} else {
    if (Test-Path (Split-Path $MountRoot -Parent)) {
        Output-Text -Filepath $LogFile -Text "Creating mountpoint root $MountRoot" -Colour Green
        New-Item $MountRoot -ItemType Directory
    } else {
        Output-Text -Filepath $LogFile -Text "Parent path for Mountpoint root ($(Split-Path $MountRoot -Parent)) doesn't exist - unable to continue" -Colour Yellow
        break 
    }
}

$RawDBDisks = @(Get-Disk | Where-Object {($_.PartitionStyle -eq "Raw") -and ($_.Size -eq $DBDiskSize)})
Write-Debug $RawDBDisks.Count
Write-Debug $DBList.Count


if ($RawDBDisks.Count -ne $DBList.Count) {
    
    Write-Error "`nNumber of Raw disks ($($RawDBDisks.Count)) doesn't match number of mountpoints to be configured ($($DBList.Count))`nExiting script"
    Break
}

$DiskCount = $RawDBDisks.Count

Output-Text -FilePath $LogFile -Text "Configuring $DiskCount database mountpoints in $MountRoot" -Colour Cyan

for ($i=0;$i -lt $DiskCount;$i++) {
    $DiskNumber = $RawDBDisks[$i].Number
    $DAGName = $DBList[$i]

    #Write-Debug 
    Write-Debug $DAGName

    Output-Text -Filepath $LogFile -Text ""
    Output-Text -Filepath $LogFile -Text "Configuring mountpoint $MountRoot\$DAGName using disk $DiskNumber" -Colour Cyan
    
    $DiskNumber = $RawDBDisks[$i].Number
    $DAGName = $DBList[$i]

        $MountPoint = Test-Path "$MountRoot\$DAGName"
    If ($MountPoint) {
        $MountPoint = Get-Item "$MountRoot\$DAGName"
        If ($MountPoint -isnot [System.IO.DirectoryInfo]) {
            Output-Text -Filepath $LogFile -Text "An item with path $MountPoint already exists, but is not a Directory" -Colour Red
            Output-Text -Filepath $LogFile -Text "Skipping $DAGName" -Colour Red
            Continue
        } elseif ((Get-ChildItem $MountPoint).count -ne 0) {
            Output-Text -Filepath $LogFile -Text "Directory $MountPoint already exists, but is not empty" -Colour Red
            Output-Text -Filepath $LogFile -Text "Skipping $DAGName" -Colour Red
            Continue
        }
    } else {
        New-Item $MountRoot\$DAGName -ItemType Directory
    }

    Initialize-Disk -Number $DiskNumber -PartitionStyle GPT 
    $Partition = New-Partition -DiskNumber $DiskNumber -UseMaximumSize 
    Start-Sleep -Seconds $Delay
    $Partition |Format-Volume -FileSystem NTFS -NewFileSystemLabel $DAGName -Confirm:$false 

    $Partition |Add-PartitionAccessPath -AccessPath "$MountRoot\$DAGName" |Out-Null
    If ((Get-Partition).AccessPaths -contains "$MountRoot\$DAGName\") {
        Output-Text -Filepath $LogFile -Text "Mountpoint for DAG $DAGName created successfully at $MountRoot\$DAGName" -Colour Green
    } else {
        Output-Text -Filepath $LogFile -Text "Failure creating mountpoint for DAG $DAGName" -Colour Red
    }
    #Output-Text -Filepath $LogFile -Text "Disk for DB $DAGName created and mounted at $MountRoot\$DAGName" -Colour Green


    #>
}

Clear-Variable Partition -ErrorAction SilentlyContinue
#***********RESTORE DISK***********

Output-Text -Filepath $LogFile -Text ""
Output-Text -Filepath $LogFile -Text "Initialising Restore Volume" -Colour Cyan


$MountPoint = Test-Path "$MountRoot\Restore"

If ($MountPoint) {
    $MountPoint = Get-Item "$MountRoot\Restore"
    If ($MountPoint -isnot [System.IO.DirectoryInfo]) {
        Output-Text -Filepath $LogFile -Text "An item with path $MountPoint already exists, but is not a Directory" -Colour Red
        Output-Text -Filepath $LogFile -Text "Skipping Restore volume" -Colour Red
        Continue
    } elseif ((Get-ChildItem $MountPoint).count -ne 0) {
        Output-Text -Filepath $LogFile -Text "Directory $MountPoint already exists, but is not empty" -Colour Red
        Output-Text -Filepath $LogFile -Text "Skipping Restore volume" -Colour Red
        Continue
    }
} else {
    New-Item "$MountRoot\Restore" -ItemType Directory
}

#CREATE MOUNTPOINT PATH
$RestoreDisk = Get-Disk | Where-Object {($_.PartitionStyle -eq "Raw") -and ($_.Size -eq $RestoreDiskSize)}
Output-Text -Filepath $LogFile -Text "Creating Restore disk using Disk $($RestoreDisk.Number)" -Colour Cyan

Initialize-Disk -Number $RestoreDisk.Number -PartitionStyle GPT
$Partition = New-Partition -DiskNumber $RestoreDisk.Number -UseMaximumSize 
Start-Sleep -Seconds $Delay
$Partition |Format-Volume -FileSystem NTFS -NewFileSystemLabel "Restore" -Confirm:$false

$Partition | Add-PartitionAccessPath -AccessPath "$MountRoot\Restore" 

If ((Get-Partition).AccessPaths -contains "$MountRoot\Restore\") {
    Output-Text -Filepath $LogFile -Text "Mountpoint for Restore volume created successfully at $MountRoot\Restore" -Colour Green
} else {
    Output-Text -Filepath $LogFile -Text "Failure creating mountpoint for Restore volume" -Colour Red
}
#Output-Text -Filepath $LogFile -Text "Restore disk completed" -Colour Green

Write-Host -ForegroundColor DarkBlue "`n`nDisk configuration completed, please review the output at $LogFile" -BackgroundColor White
Write-Host -ForegroundColor Green "`n`nNow please modify the default Defrag Scheduled Task to remove all Exchange Volumes"
Invoke-Expression "c:\windows\system32\dfrgui.exe"